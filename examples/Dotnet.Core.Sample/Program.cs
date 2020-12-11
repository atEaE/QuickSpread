using Dotnet.Core.Sample.Command.Excel;
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

            Console.WriteLine(@"
 [a] : Primitive array sample.
 [c] : Class array sample.
 [e] : End.
");
            var input = Console.ReadLine();

            var endFlag = false;
            while (!endFlag)
            {
                switch (input)
                {
                    case PRIMITIVE_ARRAY_SAMPLE_KEY:
                        new PrimitiveArray().Execute();
                        endFlag = true;
                        break;
                    case CLASS_ARRAY_SAMPLE_KEY:
                        new ClassArray().Execute();
                        endFlag = true;
                        break;
                    case END_KEY:
                        endFlag = true;
                        break;
                    default:
                        Console.WriteLine("An untargeted command was entered.");
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
