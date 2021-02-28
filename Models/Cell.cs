using System;
namespace ExcelMerger.Models
{
    public class Cell
    {
        public Type DataType { get; set; }

        public object Value { get; set; }

        public object Address { get; set; }
    }
}
