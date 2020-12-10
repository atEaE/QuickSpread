using Dotnet.Core.Sample.Base;
using System;

namespace Dotnet.Core.Sample.Command
{
    /// <summary>
    /// Primitive array sample command.
    /// </summary>
    public class PrimitiveArray : ICommand
    {
        /// <summary>
        /// Execute the command.
        /// </summary>
        public void Execute()
        {
            Console.WriteLine("Primitive array");
        }
    }
}
