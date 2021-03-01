using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelMerger.API;
using ExcelMerger.Exceptions;
using ExcelMerger.Models;
using ExcelMerger.Utilities;
using Google.Apis.Drive.v3.Data;

namespace ExcelMerger
{
    public class Merger : IDisposable
    {
        private GoogleDriveAPI driveApi;
        private GoogleSheetsAPI sheetsApi;
        public Merger()
        {
            driveApi = new GoogleDriveAPI();
            sheetsApi = new GoogleSheetsAPI();
        }

        public async Task RunAsync()
        {
            try
            {
                var files = await driveApi.ListFilesAsync();
                var mergedFileData = await MergeSheetsAsync(files);
                await CreateFileInFolder(mergedFileData);
                Console.WriteLine("Success...");
            }
            catch (CancelledException cancelledException)
            {
                Console.WriteLine(cancelledException.Message);
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error occured:");
                Console.WriteLine(ex);
            }
        }

        private async Task CreateFileInFolder(FileData fileData)
        {
            var newSpreadSheetId = await sheetsApi.GenerateSheetsFile(fileData);
            await driveApi.MoveFileToFolder(newSpreadSheetId, GetFolderInput(), fileData.FileName);
        }

        private async Task<FileData> MergeSheetsAsync(IList<File> files)
        {
            var sheetData = SetSheetData();
            List<FileData> fileDataList = new List<FileData>();
            foreach (var file in files)
            {
                // validate sheet data for file


                // get file data
                var fileData = new FileData
                {
                    FileName = file.Name,
                    Cells = await sheetsApi.ReadSheet(file.Id, sheetData)
                };

                fileDataList.Add(fileData);
            }

            // merge data into one file
            return MergeFileData(fileDataList);
        }

        private FileData MergeFileData(List<FileData> fileDataList)
        {
            FileData mergedFileData = new FileData
            {
                FileName = "Merged File " + DateTimeOffset.Now.Date.ToShortDateString()
            };
            mergedFileData.Cells = new List<Cell>();
            var firstFile = fileDataList.First();
            fileDataList.RemoveAt(0);
            foreach (var cell in firstFile.Cells)
            {
                if (!cell.IsNumber)
                {
                    if (!fileDataList.SelectMany(fd => fd.Cells).Where(c => c.Row == cell.Row && c.Column == cell.Column).All(c => c.Value == cell.Value))
                    {
                        // add some warning
                    }
                }
                else
                {
                    foreach (var value in fileDataList.SelectMany(fd => fd.Cells).Where(c => c.Row == cell.Row && c.Column == cell.Column).Select(c => c.Value).Cast<double>())
                    {
                        cell.Value += value;
                    }
                }
                mergedFileData.Cells.Add(cell);
            }

            return mergedFileData;
        }

        private SheetSettings SetSheetData()
        {
            Console.WriteLine("Enter Sheet Name:");
            string sheetName = Console.ReadLine();
            Console.WriteLine("Enter Range:");
            string range = Console.ReadLine();

            return new SheetSettings(sheetName, range);
        }
        private string GetFolderInput()
        {
            Console.WriteLine("Enter folder where the merged file will be stored:");
            return Console.ReadLine();
        }

        public void Dispose()
        {
            if (driveApi != null)
            {
                driveApi.Dispose();
                driveApi = null;
            }
            if (sheetsApi != null)
            {
                sheetsApi.Dispose();
                sheetsApi = null;
            }
        }
    }
}