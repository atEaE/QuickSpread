using Dotnet.Core.Sample.Base;

namespace Dotnet.Core.Sample.Command.Excel
{
    /// <summary>
    /// Excel command builder.
    /// </summary>
    public class ExcelCommandBuilder : ICommandBuilder
    {
        /// <summary>
        /// Generates a command for a class array.
        /// </summary>
        /// <returns>ClassArrayCommand</returns>
        public CommandBase BuildClassArrayCommand()
        {
            return new ClassArray();
        }

        /// <summary>
        /// Generates commands for primitive type arrays.
        /// </summary>
        /// <returns>PrimitiveArrayCommand</returns>
        public CommandBase BuildPrimitiveArrayCommand()
        {
            return new PrimitiveArray();
        }
    }
}
