using Dotnet.Framework.Sample.Base;
using System;

namespace Dotnet.Framework.Sample.Command.GoogleSpreadSheet
{
    /// <summary>
    /// Google spreadsheet command builder.
    /// </summary>
    public class GoogleSpreadSheetCommandBuilder : ICommandBuilder
    {
        /// <summary>
        /// Generates a command for a class array.
        /// </summary>
        /// <returns>ClassArrayCommand</returns>
        public CommandBase BuildClassArrayCommand()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Generates commands for primitive type arrays.
        /// </summary>
        /// <returns>PrimitiveArrayCommand</returns>
        public CommandBase BuildPrimitiveArrayCommand()
        {
            throw new NotImplementedException();
        }
    }
}
