using System.Collections.Generic;
using Google.Apis.Drive.v3.Data;

namespace ExcelMerger.Models
{
    public class FileData
    {
        public string FileName { get; set; }

        public List<Cell> Cells { get; set; }
    }
}