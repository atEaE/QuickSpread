using QuickSpread.Client;
using QuickSpread.Client.Excel;
using Xunit;
using System.Collections.Generic;

namespace Test_QuickSpread.Client.Excel
{
    public class Test_ExcelSpreadSheetQuickClient
    {
        [Fact]
        public void TestCreateExcelSpreadSheetQuickClient()
        {
            var client = new ClientBuilder().Build(ExcelSpreadSheetSettings.Default(), @"C:\Users\EtaAoki\Desktop\新しいフォルダー\sample.xlsx");
            var coll = new List<Sample>()
            {
                new Sample(){ Name = "hoge", Age = 21},
                new Sample(){ Name = "huga", Age = 22},
                new Sample(){ Name = "kusa", Age = 23},
            };
            client.Export(new string[] { "hogehoe", "hugahuga", "piyopiyo"}, new ExcelSpreadSheetOptions() { StartRowIndex = 2, StartColumnIndex = 3 });
        }


        public class Sample
        {
            public string Name;
            public int Age;
        }
    }
}
