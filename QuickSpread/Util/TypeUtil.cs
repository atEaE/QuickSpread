using System.Runtime.CompilerServices;
using System;

[assembly: InternalsVisibleTo("Test_QuickSpread")]
namespace QuickSpread.Util
{
    /// <summary>
    /// Type utility class.
    /// </summary>
    internal static class TypeUtil
    {
        /// <summary>
        /// Determines if the argument is a primitive type.
        /// </summary>
        /// <param name="type">any types.</param>
        /// <returns>true is primitive.</returns>
        public static bool IsPrimitive(Type type)
        {
            if (type.IsPrimitive)
            {
                return true;
            }
            
            if (type == typeof(string))
            {
                return true;
            }

            return false;
        }
    }
}
