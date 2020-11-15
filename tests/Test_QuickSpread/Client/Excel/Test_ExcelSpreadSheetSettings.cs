using QuickSpread.Client.Excel;
using System;
using Xunit;

namespace Test_QuickSpread.Client.Excel
{
    public class Test_ExcelSpreadSheetSettings
    {
        [Fact]
        public void TestCreateInstance()
        {
            // setup
            var settings = new ExcelSpreadSheetSettings();

            settings.SheetName.IsNull();
        }

        [Fact]
        public void TestCreateDefaultInstance()
        {
            // setup
            var settings = ExcelSpreadSheetSettings.Default();

            settings.SheetName.Is("Sheet1");
        }

        [Fact]
        public void TestCheckValidate_SheetName()
        {
            // setup
            var settings = ExcelSpreadSheetSettings.Default();
            settings.Validate();

            // string.Emtpy
            settings.SheetName = string.Empty;
            Assert.Throws<ApplicationException>(() => { settings.Validate(); }).Message.Is("The sheet name has not been set.");

            // ""
            settings.SheetName = "";
            Assert.Throws<ApplicationException>(() => { settings.Validate(); }).Message.Is("The sheet name has not been set.");

            // null
            settings.SheetName = null;
            Assert.Throws<ApplicationException>(() => { settings.Validate(); }).Message.Is("The sheet name has not been set.");

            // brank 
            settings.SheetName = " ";
            Assert.Throws<ApplicationException>(() => { settings.Validate(); }).Message.Is("The sheet name has not been set.");
        }
    }
}
