using QuickSpread.Client;
using QuickSpread.Client.GoogleSpreadSheet;
using Xunit;
using System.Collections.Generic;

namespace Test_QuickSpread
{
    public class Test_ClientBuilder
    {
        [Fact]
        public void TestCreateGoogleSpreadQuickClient()
        {
            var client = new ClientBuilder().Build("");
            client.IsInstanceOf<GoogleSpreadQuickClient>();
            client.Export(exportCollections: new List<Sample>() { new Sample(){ Name = "hoge", Piyo = 1 } });
        }



        public class Sample
        {
            public string Name;
            public int Piyo;
        }
    }


}
