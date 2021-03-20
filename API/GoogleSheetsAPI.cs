
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelMerger.Models;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace ExcelMerger.API
{
    public class GoogleSheetsAPI : GoogleAPI
    {
        private SheetsService sheetsService;
        public async Task<List<Cell>> ReadSheet(string fileId, SheetSettings sheetData)
        {
            if (sheetsService == null)
            {
                sheetsService = new SheetsService(GetInitializer());
            }
            var cells = new List<Cell>();
            // Define request parameters.
            string spreadsheetId = fileId;
            string range = sheetData.SheetQuery;
            SpreadsheetsResource.ValuesResource.GetRequest request = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = await request.ExecuteAsync();
            IList<IList<object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                for (int rowIndex = 0; rowIndex < values.Count; rowIndex++)
                {
                    for (int cellIndex = 0; cellIndex < values[rowIndex].Count; cellIndex++)
                    {
                        cells.Add(new Cell(values[rowIndex][cellIndex] as string, rowIndex, cellIndex));
                    }
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            return cells;
        }

        public async Task<string> GenerateSheetsFile(FileData fileData)
        {
            if (sheetsService == null)
            {
                sheetsService = new SheetsService(GetInitializer());
            }

            Spreadsheet body = CreateEmptySpreadsheet();
            var createFileRequest = sheetsService.Spreadsheets.Create(body);
            var response = await createFileRequest.ExecuteAsync();
            await InsertSheetData(response.SpreadsheetId, response.Sheets.First().Properties.SheetId, fileData);
            return response.SpreadsheetId;
        }

        private async Task InsertSheetData(string spreadsheetId, int? sheetId, FileData fileData)
        {
            string range = "A1:Z";
            var body = new ValueRange();
            body.MajorDimension = "ROWS";
            body.Range = range;
            body.Values = CreateRows(fileData);
            var request = sheetsService.Spreadsheets.Values.Append(body, spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var response = await request.ExecuteAsync();
        }

        private IList<IList<object>> CreateRows(FileData fileData)
        {
            var rows = new List<IList<object>>();
            foreach (var row in fileData.Cells.GroupBy(c => c.Row))
            {
                var rowData = new List<object>();
                foreach (var cell in row.OrderBy(c => c.Column))
                {
                    rowData.Add(cell.Value);
                    //rowData.Add(GetExtendedValueFromCell(cell));
                }
                rows.Add(rowData);
            }
            return rows;
        }

        private Spreadsheet CreateEmptySpreadsheet()
        {
            Spreadsheet spreadSheet = new Spreadsheet
            {
                Sheets = new List<Sheet>() { new Sheet { Properties = new SheetProperties { Title = "Merged" } } },
                Properties = new SpreadsheetProperties { Title = "Merged Sheet" }
            };
            return spreadSheet;
        }

        private ExtendedValue GetExtendedValueFromCell(Cell cell)
        {
            var value = new ExtendedValue();
            if (cell.IsNumber)
            {
                value.NumberValue = cell.Value;
            }
            value.StringValue = cell.Value.ToString();
            return value;
        }

        public override void Dispose()
        {
            if (sheetsService != null)
            {
                sheetsService.Dispose();
                sheetsService = null;
            }
        }
    }
}