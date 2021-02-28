using System;
using System.Threading.Tasks;
using ExcelMerger.API;
namespace ExcelMerger
{
    class Program
    {
        static void Main(string[] args)
        {
            var pr = new Program();
            pr.Run().Wait();
            Console.ReadKey();
        }

        private async Task Run()
        {
            await RunDrive();
            //await RunSheets();
        }

        private async Task RunDrive()
        {
            var driveApi = new GoogleDriveAPI();
            await driveApi.ListFiles();
        }

        private async Task RunSheets()
        {
            var sheetApi = new GoogleSheetsAPI();
            await sheetApi.ReadSheet();
        }
    }
}
