using System;

namespace Dotnet.Core.Sample.Base
{
    /// <summary>
    /// Command interface
    /// </summary>
    public abstract class CommandBase
    {
        /// <summary>
        /// yes key.
        /// </summary>
        private const string CHECK_YES_KEY = "y";

        /// <summary>
        /// no key.
        /// </summary>
        private const string CHECK_NO_KEY = "n";

        /// <summary>
        /// Output file path.
        /// </summary>
        protected abstract string OutputFilePath { get; }

        /// <summary>
        /// Execute the command.
        /// </summary>
        public void Execute()
        {
            var checkKey = PreExecute();
            Console.WriteLine();

            switch (checkKey)
            {
                case CHECK_YES_KEY:
                    OnExecute();
                    break;
                case CHECK_NO_KEY:
                default:
                    break;
            }
        }

        /// <summary>
        /// Pre execute function.
        /// </summary>
        protected string PreExecute()
        {
            Console.WriteLine($"Output the file to\r\n ⇒ {OutputFilePath} \r\nAre you sure?");
            Console.Write("[y] or [n] : ");
            return Console.ReadLine();
        }

        /// <summary>
        /// On Execute the command.
        /// </summary>
        protected abstract void OnExecute();
    }
}
