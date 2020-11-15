using System;

namespace QuickSpread.Client.Excel
{
    /// <summary>
    /// Excel spread sheet options.
    /// </summary>
    public class ExcelSpreadSheetSettings : ISpreadSheetSettings
    {
        /// <summary>
        /// Excel sheetname.;
        /// </summary>

        public string SheetName { get; set; }

        /// <summary>
        /// Gets the default value of Excel settings.
        /// </summary>
        /// <returns>default settings.</returns>
        public static ExcelSpreadSheetSettings Default()
        {
            return new ExcelSpreadSheetSettings()
            {
                SheetName = "Sheet1",
            };
        }

        /// <summary>
        /// Evaluate the effectiveness.
        /// </summary>
        public void Validate()
        {
            if (string.IsNullOrEmpty(SheetName) || string.IsNullOrWhiteSpace(SheetName))
                throw new ApplicationException("The sheet name has not been set.");
        }
    }
}
