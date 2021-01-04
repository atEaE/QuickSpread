using System;

namespace QuickSpread.Client.Excel
{
    /// <summary>
    /// Set where the spreadsheet header information will be read from.
    /// </summary>
    public enum ReadHeaderInfo
    {
        /// <summary>
        /// Reads the header information from the property.
        /// </summary>
        Property,

        /// <summary>
        /// Reads the header information from the field.
        /// </summary>
        Field,

        /// <summary>
        /// Reads the header information from the property and field.
        /// </summary>
        PropertyAndField,
    }

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
        /// Set where the spreadsheet header information will be read from.
        /// </summary>
        public ReadHeaderInfo ReadHeaderInfo { get; set; }

        /// <summary>
        /// Gets the default value of Excel settings.
        /// </summary>
        /// <returns>default settings.</returns>
        public static ExcelSpreadSheetSettings Default()
        {
            return new ExcelSpreadSheetSettings()
            {
                SheetName = "Sheet1",
                ReadHeaderInfo = ReadHeaderInfo.Property,
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
