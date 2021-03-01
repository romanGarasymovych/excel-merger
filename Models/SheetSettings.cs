using System;
namespace ExcelMerger.Models
{
    public class SheetSettings
    {
        private readonly string sheetName;
        private readonly string range;

        public SheetSettings(string sheetName, string range)
        {
            this.sheetName = sheetName;
            this.range = range;
        }

        public string SheetQuery => $"{sheetName}!{range}";
    }
}
