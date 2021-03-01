using System;
namespace ExcelMerger.Models
{
    public class Cell
    {
        public bool IsNumber { get; private set; }

        public dynamic Value { get; set; }

        public int Row { get; set; }

        public int Column { get; set; }

        public Cell(string value, int row, int column)
        {
            IsNumber = double.TryParse(value, out double numericValue);
            if (IsNumber)
            {
                Value = numericValue;
            }
            else
            {
                Value = value;
            }
            Row = row;
            Column = column;
        }

    }
}
