namespace QuickSpread.Client.Excel
{
    /// <summary>
    /// Datetime cell format
    /// </summary>
    public struct DateTimeFormat
    {
        /// <summary>
        /// format.
        /// </summary>
        internal string Format { get; set; }
    }

    /// <summary>
    /// Excel spread sheet options.
    /// </summary>
    public class ExcelSpreadSheetOptions : ISpreadSheetOptions
    {
        /// <summary>
        /// Output start line index
        /// </summary>
        public uint StartRowIndex { get; set; } = 0;

        /// <summary>
        /// Output start column index
        /// </summary>
        public uint StartColumnIndex { get; set; } = 0;

        /// <summary>
        /// Cell format for Datetime.
        /// </summary>
        public DateTimeFormat DataTimeCellFormat { get; set; } = new DateTimeFormat { Format = "yyyy/MM/dd HH:mm:ss" };

        /// <summary>
        /// Evaluate the effectiveness.
        /// </summary>
        public void Validate()
        { }
    }
}
