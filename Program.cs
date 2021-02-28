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
            var pr = new Program();
            pr.Run().Wait();
            Console.ReadKey();
        }

        private async Task Run()
        {
            var files = await RunDrive();
            await RunSheets(files);
        }

        private async Task<IList<File>> RunDrive()
        {
            var driveApi = new GoogleDriveAPI();
            return await driveApi.ListFiles();
        }

        private async Task RunSheets(IList<File> files)
        {
            var sheetApi = new GoogleSheetsAPI();
            await sheetApi.ReadSheet();
        }
    }
}
