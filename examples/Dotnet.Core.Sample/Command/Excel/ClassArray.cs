using AutoFixture;
using Dotnet.Core.Sample.Base;
using Dotnet.Core.Sample.Model;
using QuickSpread.Client;
using QuickSpread.Client.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Dotnet.Core.Sample.Command.Excel
{
    /// <summary>
    /// Class array sample command.
    /// </summary>
    public class ClassArray : CommandBase
    {
        /// <summary>
        /// Output base directory.
        /// </summary>
        protected string BaseOutputDirectory => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excel_example_dotnet_core");

        /// <summary>
        /// Output file path.
        /// </summary>
        protected override string OutputFilePath => Path.Combine(BaseOutputDirectory, "class_sample.xlsx");

        /// <summary>
        /// On Execute the command.
        /// </summary>
        protected override void OnExecute()
        {
            
            var client = new ClientBuilder().Build(ExcelSpreadSheetSettings.Default(), OutputFilePath);
            var list = createOutputModel();

            if (!Directory.Exists(BaseOutputDirectory))
            {
                Directory.CreateDirectory(BaseOutputDirectory);
            }
            if (File.Exists(OutputFilePath))
            {
                Console.WriteLine("The file already exists. Please check the file.");
            }
            else
            {
                client.Export(list);
            }
        }

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
