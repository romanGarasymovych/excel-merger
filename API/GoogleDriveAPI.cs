using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using ExcelMerger.Exceptions;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;

namespace ExcelMerger.API
{
    public class GoogleDriveAPI : GoogleAPI
    {
        private DriveService driveService;

        public async Task<IList<File>> ListFilesAsync()
        {
            if (driveService == null)
            {
                driveService = new DriveService(GetInitializer());
            }

            // Define parameters of request.
            FilesResource.ListRequest listRequest = driveService.Files.List();
            string folderId = GetFolderInput();
            listRequest.Q = $"'{folderId}' in parents";
            listRequest.PageSize = 100;
            listRequest.Fields = "nextPageToken, files(id, name)";

            // List files.
            FileList fileList = await listRequest.ExecuteAsync();
            var files = fileList.Files;
            if (files == null || files.Count == 0)
            {
                Console.WriteLine("No files found. Would you like to try again? (Y/N)");
                string answer = Console.ReadLine();
                return answer switch
                {
                    "Y" => await ListFilesAsync(),
                    _ => throw new CancelledException(),
                };
            }
            Console.WriteLine($"Found {files.Count} files");
            return files;
        }

        public async Task MoveFileToFolder(string fileId, string folderId, string newFileName = null)
        {
            if (driveService == null)
            {
                driveService = new DriveService(GetInitializer());
            }
            var getFileRequest = driveService.Files.Get(fileId);
            getFileRequest.Fields = "parents";
            var file = await getFileRequest.ExecuteAsync();
            //if (!string.IsNullOrWhiteSpace(newFileName))
            //{
            //    file.Name = newFileName;
            //}
            var updateFileRequest = driveService.Files.Update(file, fileId);
            updateFileRequest.AddParents = folderId;
            updateFileRequest.RemoveParents = file.Parents.First();
            updateFileRequest.Fields = "addParents, removeParents";
            var movedFile = await updateFileRequest.ExecuteAsync();
        }

        private string GetFolderInput()
        {
            Console.WriteLine("Enter folder Id (last part of the url when folder is open. Usually, a long sequence of numbers and letters):");
            string input = Console.ReadLine();

            return input;
        }
        public override void Dispose()
        {
            if (driveService != null)
            {
                driveService.Dispose();
                driveService = null;
            }
        }
    }
}