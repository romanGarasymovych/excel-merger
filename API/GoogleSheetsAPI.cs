
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

            Spreadsheet body = CreateSpreadsheetFromFileData(fileData);
            var createFileRequest = sheetsService.Spreadsheets.Create(body);
            var response = await createFileRequest.ExecuteAsync();
            return response.SpreadsheetId;
        }

        private Spreadsheet CreateSpreadsheetFromFileData(FileData fileData)
        {
            Spreadsheet spreadSheet = new Spreadsheet();
            spreadSheet.Sheets = new List<Sheet>();
            spreadSheet.Sheets.Add(CreateSheetFromFileData(fileData));
            return spreadSheet;
        }

        private Sheet CreateSheetFromFileData(FileData fileData)
        {
            Sheet sheet = new Sheet
            {
                Data = new List<GridData>()
            };
            GridData data = new GridData
            {
                StartColumn = 0,
                StartRow = 0
            };

            int maxRow = fileData.Cells.Max(c => c.Row);
            int maxCol = fileData.Cells.Max(c => c.Column);
            for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++)
            {
                RowData row = new RowData
                {
                    Values = new List<CellData>()
                };
                foreach (var cell in fileData.Cells.Where(c => c.Row == rowIndex).OrderBy(c => c.Column))
                {
                    row.Values.Add(new CellData() { UserEnteredValue = GetExtendedValueFromCell(cell) });
                }
            }

            sheet.Data.Add(data);

            return sheet;
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