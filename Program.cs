using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelMerger.API;
using Google.Apis.Drive.v3.Data;

namespace ExcelMerger
{
    class Program
    {
        static void Main(string[] args)
        {
            var merger = new Merger();
            merger.RunAsync().Wait();
            Console.ReadKey();
        }
    }
}
