using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelMerger.API
{
    public abstract class GoogleAPI : IDisposable
    {
        public const string ApplicationName = "Excel Merger";

        protected string[] GetScopes() => new string[]
        {
            SheetsService.Scope.SpreadsheetsReadonly,
            DriveService.Scope.Drive
        };

        protected BaseClientService.Initializer GetInitializer()
        {
            var credentials = GetCredentials();
            return new BaseClientService.Initializer
            {
                ApplicationName = ApplicationName,
                HttpClientInitializer = credentials
            };
        }

        private UserCredential GetCredentials()
        {
            using var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read);
            // The file token.json stores the user's access and refresh tokens, and is created
            // automatically when the authorization flow completes for the first time.
            string credPath = "token.json";
            UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.Load(stream).Secrets,
                GetScopes(),
                "user",
                CancellationToken.None,
                new FileDataStore(credPath, true)).Result;
            Console.WriteLine("Credential file saved to: " + credPath);
            return credential;
        }

        public abstract void Dispose();
    }
}
