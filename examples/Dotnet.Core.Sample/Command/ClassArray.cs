using Dotnet.Core.Sample.Base;
using System;

namespace Dotnet.Core.Sample.Command
{
    /// <summary>
    /// Class array sample command.
    /// </summary>
    public class ClassArray : ICommand
    {
        /// <summary>
        /// Execute the command.
        /// </summary>
        public void Execute()
        {
            Console.WriteLine("Class array");
        }
    }
}
