using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using QuickSpread.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace QuickSpread.Client.Excel
{
    /// <summary>
    /// Extend the Builder for ExcelSpreadSheetQuickClient.
    /// </summary>
    public static class ClientBuilderExtensions
    {
        /// <summary>
        /// Generate a ExcelSpreadSheetQuickClient.
        /// </summary>
        public static IQuickClient Build(this ClientBuilder _, ExcelSpreadSheetSettings settings, string filePath)
        {
            return new ExcelSpreadSheetQuickClient(settings: settings, filePath: filePath);
        }
    }

    public class ExcelSpreadSheetQuickClient : IQuickClient
    {
        /// <summary>
        /// Microsoft Excel Format : xls (excel 97-2003)
        /// </summary>
        private const string HSSF_EXTENSION = ".xls";

        /// <summary>
        /// Office Open XML Workbook format : xlsx (excel 2007 -)
        /// </summary>
        private const string XSSF_EXTENSION = ".xlsx";

        /// <summary>
        /// filePath
        /// </summary>
        private string filePath;

        /// <summary>
        /// excel settings
        /// </summary>
        private ExcelSpreadSheetSettings settings;

        /// <summary>
        /// Create new instance.
        /// </summary>
        /// <param name="settings">excel settings</param>
        /// <param name="filePath">filePath</param>
        internal ExcelSpreadSheetQuickClient(ExcelSpreadSheetSettings settings, string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException(nameof(filePath));

            settings.Validate();

            this.settings = settings;
            this.filePath = filePath;
        }

        /// <summary>
        /// Export to a specified spreadsheet.
        /// </summary>
        /// <typeparam name="T">POCO with no internal List, Array or Class</typeparam>
        /// <param name="exportCollections">export collections</param>
        /// <param name="options">excel options.</param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="IOException"></exception>
        public void Export<T>(IList<T> exportCollections, ISpreadSheetOptions options = null)
        {
            if (exportCollections == null)
                throw new ArgumentNullException(nameof(exportCollections));

            if (!exportCollections.Any())
                return;

            if (File.Exists(filePath))
                throw new IOException($"{filePath} already exists.");

            ExcelSpreadSheetOptions excelOpt = null;
            if (options != null)
            {
                if (!(options is ExcelSpreadSheetOptions))
                    throw new ArgumentException($"The only Options that this class can receive are {nameof(ExcelSpreadSheetOptions)}.");

                excelOpt = options as ExcelSpreadSheetOptions;
            }

            IWorkbook book;
            var extension = Path.GetExtension(filePath);
            if (extension == HSSF_EXTENSION)
            {
                book = new HSSFWorkbook();
            }
            else if (extension == XSSF_EXTENSION)
            {
                book = new XSSFWorkbook();
            }
            else
            {
                throw new ApplicationException("CreateNewBook: invalid extension");
            }

            var sheet = book.CreateSheet(settings.SheetName);

            if (TypeUtil.IsPrimitive(typeof(T)))
            {
                PrimitiveTypeExport(sheet, exportCollections, excelOpt);
            }
            else
            {
                ClassTypeExport(sheet, exportCollections, excelOpt);
            }

            using (var fs = new FileStream(filePath, FileMode.Create))
            {
                book.Write(fs);
            }
        }

        /// <summary>
        /// Outputs primitive type.
        /// </summary>
        /// <typeparam name="T">any primitive types.</typeparam>
        /// <param name="sheet">excel sheet.</param>
        /// <param name="exportCollections">export any collections.</param>
        /// <param name="options">excel sheet options.</param>
        protected virtual void PrimitiveTypeExport<T>(ISheet sheet, IList<T> exportCollections, ExcelSpreadSheetOptions options)
        {
            var rowIndex = 0;
            var columnIndex = 0;

            if (options != null)
            {
                rowIndex = (int)options.StartRowIndex;
                columnIndex = (int)options.StartColumnIndex;
            }

            foreach (var value in exportCollections)
            {
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
                cell.SetCellValue(value.ToString());
                rowIndex++;
            }
        }

        /// <summary>
        /// Outputs class type.
        /// </summary>
        /// <typeparam name="T">any class types.</typeparam>
        /// <param name="sheet">excel sheet.</param>
        /// <param name="exportCollections">export any collections.</param>
        /// <param name="options">excel sheet options.</param>
        protected virtual void ClassTypeExport<T>(ISheet sheet, IList<T> exportCollections, ExcelSpreadSheetOptions options) 
        {
            var rowIndex = 0;
            var columnIndex = 0;

            if (options != null)
            {
                rowIndex = (int)options.StartRowIndex;
                columnIndex = (int)options.StartColumnIndex;
            }

            var hColIndex = 0;
            foreach (var prop in typeof(T).GetProperties())
            {
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                var cell = row.GetCell(hColIndex) ?? row.CreateCell(hColIndex);
                cell.SetCellValue(prop.Name);
                hColIndex++;
            }

            rowIndex = rowIndex + 1;
            foreach (var value in exportCollections)
            {
                var colIndex = columnIndex;
                foreach(var prop in typeof(T).GetProperties())
                {
                    var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                    var cell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
                    cell.SetCellValue(prop.GetValue(value).ToString());
                    colIndex++;
                }
                rowIndex++;
            }
        }
    }
}
