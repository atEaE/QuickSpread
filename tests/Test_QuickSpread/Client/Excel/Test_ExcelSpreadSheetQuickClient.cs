using QuickSpread.Client;
using QuickSpread.Client.Excel;
using System;
using Xunit;

namespace Test_QuickSpread.Client.Excel
{
    public class Test_ExcelSpreadSheetQuickClient
    {
        [Fact]
        public void TestCreateExcelSpreadSheetQuickClient_SettingsError_SheetName()
        {
            // setup
            var settings = new ExcelSpreadSheetSettings();

            // string.Empty
            Assert.Throws<ApplicationException>(() =>
            {
                settings.SheetName = string.Empty;
                var client = new ClientBuilder().Build(settings, "./sample.xlsx");
            }).Message.Is("The sheet name has not been set.");

            // ""
            Assert.Throws<ApplicationException>(() =>
            {
                settings.SheetName = "";
                var client = new ClientBuilder().Build(settings, "./sample.xlsx");
            }).Message.Is("The sheet name has not been set.");

            // null
            Assert.Throws<ApplicationException>(() =>
            {
                settings.SheetName = null;
                var client = new ClientBuilder().Build(settings, "./sample.xlsx");
            }).Message.Is("The sheet name has not been set.");

            // brank
            Assert.Throws<ApplicationException>(() =>
            {
                settings.SheetName = " ";
                var client = new ClientBuilder().Build(settings, "./sample.xlsx");
            }).Message.Is("The sheet name has not been set.");
        }

        [Fact]
        public void TestCreateExcelSpreadSheetQuickClient_FilepathError()
        {
            // string.empty
            Assert.Throws<ArgumentNullException>(() =>
            {
                var client = new ClientBuilder().Build(ExcelSpreadSheetSettings.Default(), string.Empty);
            }).Message.Is("Value cannot be null. (Parameter 'filePath')");

            // ""
            Assert.Throws<ArgumentNullException>(() =>
            {
                var client = new ClientBuilder().Build(ExcelSpreadSheetSettings.Default(), "");
            }).Message.Is("Value cannot be null. (Parameter 'filePath')");

            // null
            Assert.Throws<ArgumentNullException>(() =>
            {
                var client = new ClientBuilder().Build(ExcelSpreadSheetSettings.Default(), null);
            }).Message.Is("Value cannot be null. (Parameter 'filePath')");

            // brank
            Assert.Throws<ArgumentNullException>(() =>
            {
                var client = new ClientBuilder().Build(ExcelSpreadSheetSettings.Default(), " ");
            }).Message.Is("Value cannot be null. (Parameter 'filePath')");
        }
    }
}
