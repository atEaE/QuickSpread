namespace QuickSpread.Client.Excel
{
    /// <summary>
    /// Cell format
    /// </summary>
    public struct CellFormat
    {
        /// <summary>
        /// format.
        /// </summary>
        internal string Format { get; set; }
    }

    public class DateTimeFormat
    {
        /// <summary>
        /// Datetime fomat → yyyy/MM/dd HH:mm:ss
        /// </summary>
        public static CellFormat yyyyMMdd_HHmmss = new CellFormat() { Format = "yyyy/MM/dd HH:mm:ss" };
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
        public CellFormat DataTimeCellFormat { get; set; } = DateTimeFormat.yyyyMMdd_HHmmss;

        /// <summary>
        /// Evaluate the effectiveness.
        /// </summary>
        public void Validate()
        { }
    }
}
