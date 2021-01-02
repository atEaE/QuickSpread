using Dotnet.Core.Sample.Command.Excel;
using Dotnet.Core.Sample.Command.GoogleSpreadSheet;
using Dotnet.Core.Sample.Base;
using System;
using System.Diagnostics;

namespace Dotnet.Core.Sample
{
    /// <summary>
    /// Entry class.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// excell sample key const.
        /// </summary>
        private const string EXCELL_SAMPLE_KEY = "x";

        /// <summary>
        /// Google Spread sample key const.
        /// </summary>
        private const string GOOGLESPREAD_SAMPLE_KEY = "g";

        /// <summary>
        /// primitive array sample key const.
        /// </summary>
        private const string PRIMITIVE_ARRAY_SAMPLE_KEY = "a";

        /// <summary>
        /// class array sample key const.
        /// </summary>
        private const string CLASS_ARRAY_SAMPLE_KEY = "c";

        /// <summary>
        /// end key const.
        /// </summary>
        private const string END_KEY = "e";

        /// <summary>
        /// Entry points.
        /// </summary>
        /// <param name="args">commandline arguments.</param>
        public static void Main(string[] args)
        {
            if (isDoubleStartUp())
            {
                Console.WriteLine("Since multiple processes of the same type are running, this process will be terminated.");
                return;
            }

            Console.WriteLine(@"
######################################
####      QuickSpread Sample      ####
######################################");

            Console.WriteLine($@"
 [{EXCELL_SAMPLE_KEY}] : ExcellSheet Sample.
 [{GOOGLESPREAD_SAMPLE_KEY}] : Google SpreadSheet Sample.
 [e] : End.
");
            var endFlag = false;
            ICommandBuilder builder = null;

            while (!endFlag)
            {
                var input = Console.ReadLine();
                switch (input)
                {
                    case EXCELL_SAMPLE_KEY:
                        endFlag = true;
                        builder = new ExcelCommandBuilder();
                        break;
                    case GOOGLESPREAD_SAMPLE_KEY:
                        endFlag = true;
                        builder = new GoogleSpreadSheetCommandBuilder();
                        break;
                    case END_KEY:
                        Console.WriteLine("Exit the sample application.");
                        Console.ReadKey();
                        return;
                    default:
                        Console.WriteLine("An untargeted command was entered. Enter the correct command.");
                        break;
                }
            }

            Console.WriteLine("Select the sample you want to create.");
            Console.WriteLine($@"
 [{PRIMITIVE_ARRAY_SAMPLE_KEY}] : Primitive array sample.
 [{CLASS_ARRAY_SAMPLE_KEY}] : Class array sample.
 [e] : End.
");

            endFlag = false;
            while (!endFlag)
            {
                var input = Console.ReadLine();
                switch (input)
                {
                    case PRIMITIVE_ARRAY_SAMPLE_KEY:
                        builder.BuildPrimitiveArrayCommand().Execute();
                        endFlag = true;
                        break;
                    case CLASS_ARRAY_SAMPLE_KEY:
                        builder.BuildClassArrayCommand().Execute();
                        endFlag = true;
                        break;
                    case END_KEY:
                        endFlag = true;
                        break;
                    default:
                        Console.WriteLine("An untargeted command was entered. Enter the correct command.");
                        break;
                }
            }

            Console.WriteLine("Exit the sample application.");
            Console.ReadKey();
        }

        /// <summary>
        /// Make sure that you are double booting.
        /// </summary>
        /// <returns>true : double-booting.</returns>
        private static bool isDoubleStartUp()
        {
            //If there are multiple identical processes, they are double-started.
            return Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1;
        }
    }
}
