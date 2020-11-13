using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace QuickSpread.Client.GoogleSpreadSheet
{
    /// <summary>
    /// Extend the Builder for GoogleSpreadQuickClient.
    /// </summary>
    public static class ClientBuilderExtensions
    {
        /// <summary>
        /// Generate a GoogleSpreadQuickClient.
        /// </summary>
        public static IQuickClient Build(this ClientBuilder _, string value)
        {
            return new GoogleSpreadQuickClient();
        }
    }

    public class GoogleSpreadQuickClient : IQuickClient
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Hoge App";

        /// <summary>
        /// Export to a specified spreadsheet.
        /// </summary>
        /// <param name="exportCollections">export collections</param>
        /// <exception cref="ArgumentNullException"></exception>
        public void Export<T>(IList<T> exportCollections)
        {
            if (exportCollections == null)
                throw new ArgumentNullException(nameof(exportCollections));

            if (!exportCollections.Any())
                return;


            UserCredential credential;

            using (var stream = new FileStream("C:/Workspace/credentials/credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "C:/Workspace/credentials/token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                                    GoogleClientSecrets.Load(stream).Secrets,
                                    Scopes,
                                    "user",
                                    CancellationToken.None,
                                    new FileDataStore(credPath, true)).Result;
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // sheet id.
            String spreadsheetId = "";

            // input
            var wv = new List<IList<object>>()
            {
                new List<object>{"=ROW()","Bです","日付：", DateTime.Now.ToString()}
            };
            var body = new ValueRange() { Values = wv };
            var req = service.Spreadsheets.Values.Append(body, spreadsheetId, "Sheet1!A1");
            req.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var result = req.Execute();
        }
    }
}
