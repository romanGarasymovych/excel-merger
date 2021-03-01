using System;

namespace ExcelMerger.Exceptions
{
    public class CancelledException : Exception
    {
        public override string Message => "Cancelled by user";
    }
}