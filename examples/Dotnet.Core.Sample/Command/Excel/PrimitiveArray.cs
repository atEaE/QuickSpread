using Dotnet.Core.Sample.Base;
using QuickSpread.Client;
using QuickSpread.Client.Excel;
using System;
using System.IO;
using System.Linq;

namespace Dotnet.Core.Sample.Command.Excel
{
    /// <summary>
    /// Primitive array sample command.
    /// </summary>
    public class PrimitiveArray : CommandBase
    {
        /// <summary>
        /// Output base directory.
        /// </summary>
        protected string BaseOutputDirectory => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excel_example_dotnet_core");

        /// <summary>
        /// Output file path.
        /// </summary>
        protected override string OutputFilePath => Path.Combine(BaseOutputDirectory, "primitive_sample.xlsx");

        /// <summary>
        /// On Execute the command.
        /// </summary>
        protected override void OnExecute()
        {
            var list = Enumerable.Range(0, 200).Select(i => $"Sample_Message-{i}").ToList();
            var client = new ClientBuilder().Build(ExcelSpreadSheetSettings.Default(), OutputFilePath);

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
    }
}
