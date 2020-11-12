using System;
using System.Collections.Generic;
using System.Text;

namespace QuickSpread.Client.GoogleSpreadSheet
{
    /// <summary>
    /// Extend the Builder for GoogleSpreadQuickClient.
    /// </summary>
    public static class ClientBuilderExtensions
    {
        /// <summary>
        /// Generate a GoogleSpreadQuickClient.
        /// </summary>
        public static IQuickClient Build(this ClientBuilder _, string value)
        {
            return new GoogleSpreadQuickClient();
        }
    }

    public class GoogleSpreadQuickClient : IQuickClient
    {
        public void Export()
        {
            throw new NotImplementedException();
        }
    }
}
