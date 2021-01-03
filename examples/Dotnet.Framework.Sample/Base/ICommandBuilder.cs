namespace Dotnet.Framework.Sample.Base
{
    /// <summary>
    /// Command builder interface.
    /// </summary>
    public interface ICommandBuilder
    {
        /// <summary>
        /// Generates a command for a class array.
        /// </summary>
        /// <returns>ClassArrayCommand</returns>
        CommandBase BuildClassArrayCommand();

        /// <summary>
        /// Generates commands for primitive type arrays.
        /// </summary>
        /// <returns>PrimitiveArrayCommand</returns>
        CommandBase BuildPrimitiveArrayCommand();
    }
}
