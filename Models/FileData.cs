using System.Collections.Generic;
using System.Linq;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Sheets.v4.Data;

namespace ExcelMerger.Models
{
    public class FileData
    {
        public string FileName { get; set; }

        public List<Cell> Cells { get; set; }

        public GridRange GetGridRange(int? sheetId) => new GridRange()
        {
            SheetId = sheetId,
            StartColumnIndex = Cells.Min(c => c.Column),
            StartRowIndex = Cells.Min(c => c.Row),
            EndColumnIndex = Cells.Max(c => c.Column),
            EndRowIndex = Cells.Max(c => c.Row),
        };
    }
}