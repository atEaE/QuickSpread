using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using QuickSpread.Util;
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
        /// <typeparam name="T">POCO with no internal List, Array or Class</typeparam>
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

            if (TypeUtil.IsPrimitive(typeof(T)))
            {
                PrimitiveTypeExport(service, exportCollections);
            }
            else
            {
                ClassTypeExport(service, exportCollections);
            }
        }

        protected virtual void PrimitiveTypeExport<T>(SheetsService service, IList<T> exportCollections)
        {
            // sheet id.
            String spreadsheetId = "";

            List<IList<object>> exports = exportCollections.Select(e => 
            {
                return new List<object> { e } as IList<object>;
            }).ToList();

            var body = new ValueRange() { Values = exports };
            var req = service.Spreadsheets.Values.Append(body, spreadsheetId, "Sheet1!A1");
            req.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var result = req.Execute();
        }

        protected virtual void ClassTypeExport<T>(SheetsService service, IList<T> exportCollections)
        {
            // sheet id.
            String spreadsheetId = "";

            var gType = typeof(T);
            var properties = gType.GetProperties();
            var props = gType.GetFields();

            var wv = new List<IList<object>>();

            IList<object> header = new List<object>();
            foreach(var prop in gType.GetFields())
            {
                header.Add(prop.Name);
            }
            wv.Add(header);


            foreach(var coll in exportCollections)
            {
                IList<object> values = new List<object>();
                foreach (var prop in gType.GetFields())
                {
                    values.Add(prop.GetValue(coll));
                }
                wv.Add(values);
            }

            var body = new ValueRange() { Values = wv };
            var req = service.Spreadsheets.Values.Append(body, spreadsheetId, "Sheet1!A1");
            req.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var result = req.Execute();
        }
    }
}
