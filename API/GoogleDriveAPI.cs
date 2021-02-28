using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;

namespace ExcelMerger.API
{
    public class GoogleDriveAPI : GoogleAPI
    {
        public async Task<IList<File>> ListFiles()
        {
            // Create the service.
            var service = new DriveService(GetInitializer());

            // Define parameters of request.
            FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.Q = "'16zviFe8TDxWATb9aLHucZlLZUQ6nRybv' in parents";
            listRequest.PageSize = 10;
            listRequest.Fields = "nextPageToken, files(id, name)";

            // List files.
            FileList fileList = await listRequest.ExecuteAsync();
            var files = fileList.Files;
            Console.WriteLine("Files:");
            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {
                    Console.WriteLine("{0} ({1})", file.Name, file.Id);
                }
            }
            else
            {
                Console.WriteLine("No files found.");
            }

            return files;
        }
    }
}