
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace ExcelMerger.API
{
    public class GoogleSheetsAPI : GoogleAPI
    {
        public async Task ReadSheet()
        {
            var service = new SheetsService(GetInitializer());
            // Define request parameters.
            String spreadsheetId = "1Tws-2_XHnzSlanl-Wr7aCXXKH1eeWGPuqbW0zMJj2lU";
            String range = "Sheet1!A1:E";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            ValueRange response = await request.ExecuteAsync();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}, {1}", row[0], row[1]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }
    }
}