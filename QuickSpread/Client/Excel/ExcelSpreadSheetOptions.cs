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

    /// <summary>
    /// Cell format for Datetime.
    /// </summary>
    public class DateTimeFormat
    {
        /// <summary>
        /// Datetime fomat → yyyy/MM/dd HH:mm:ss
        /// </summary>
        public static CellFormat yyyyMMdd_HHmmss = new CellFormat() { Format = "yyyy/MM/dd HH:mm:ss" };

        /// <summary>
        /// Datetime fomat → yyyy/MM/dd
        /// </summary>
        public static CellFormat yyyyMMdd = new CellFormat() { Format = "yyyy/MM/dd" };

        /// <summary>
        /// Datetime fomat → yyyy/MM/dd
        /// </summary>
        public static CellFormat yyyyMMdd_HHmm = new CellFormat() { Format = "yyyy/MM/dd HH:mm" };

        /// <summary>
        /// Generate your own format.
        /// </summary>
        /// <param name="format">string format.</param>
        /// <returns>cell format.</returns>
        public static CellFormat CreateCellDateTimeFormat(string format)
        {
            return new CellFormat() { Format = format };
        }
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
