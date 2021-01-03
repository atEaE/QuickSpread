using AutoFixture;
using Dotnet.Framework.Sample.Base;
using Dotnet.Framework.Sample.Model;
using QuickSpread.Client;
using QuickSpread.Client.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Dotnet.Framework.Sample.Command.Excel
{
    /// <summary>
    /// Class array sample command.
    /// </summary>
    public class ClassArray : CommandBase
    {
        /// <summary>
        /// Output file path.
        /// </summary>
        protected override string OutputFilePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excel_example_dotnet_framework");

        /// <summary>
        /// On Execute the command.
        /// </summary>
        protected override void OnExecute()
        {
            var clients = new List<IQuickClient>();

            var pSettings = ExcelSpreadSheetSettings.Default();
            var propClient = new ClientBuilder().Build(pSettings, Path.Combine(OutputFilePath, "prop_example.xlsx"));
            clients.Add(propClient);

            var fSettings = ExcelSpreadSheetSettings.Default();
            fSettings.ReadHeaderInfo = ReadHeaderInfo.Field;
            var fildClient = new ClientBuilder().Build(fSettings, Path.Combine(OutputFilePath, "field_example.xlsx"));
            clients.Add(fildClient);

            var pfSettings = ExcelSpreadSheetSettings.Default();
            pfSettings.ReadHeaderInfo = ReadHeaderInfo.PropertyAndField;
            var pAndFClient = new ClientBuilder().Build(pfSettings, Path.Combine(OutputFilePath, "prop_field_example.xlsx"));
            clients.Add(pAndFClient);

            var list = createOutputModel();

            if (!Directory.Exists(OutputFilePath))
            {
                Directory.CreateDirectory(OutputFilePath);
            }
            if (File.Exists(OutputFilePath))
            {
                Console.WriteLine("The file already exists. Please check the file.");
            }
            else
            {
                clients.AsParallel().ForAll(c => c.Export(list));
            }
        }

        /// <summary>
        /// Create sample model.
        /// </summary>
        /// <returns>sample model.</returns>
        private List<SampleModel> createOutputModel()
        {
            return Enumerable.Range(0, 100).Select(_ =>
            {
                var fixture = new Fixture();
                return fixture.Create<SampleModel>();
            }).ToList();
        }
    }
}
