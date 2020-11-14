namespace QuickSpread.Client.Excel
{
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
        /// Evaluate the effectiveness.
        /// </summary>
        public void Validate()
        { }
    }
}
