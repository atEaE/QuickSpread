using System;
using Xunit;
using QuickSpread.Client;
using QuickSpread.Client.GoogleSpreadSheet;

namespace Test_QuickSpread
{
    public class Test_ClientBuilder
    {
        [Fact]
        public void TestCreateGoogleSpreadQuickClient()
        {
            var client = new ClientBuilder().Build("");
            client.IsInstanceOf<GoogleSpreadQuickClient>();
        }
    }
}
