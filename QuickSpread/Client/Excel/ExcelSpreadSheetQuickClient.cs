using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Model;
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

    /// <summary>
    /// QuickClient for Excel
    /// </summary>
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
                throw new ApplicationException($@"Invalid extension has been specified. Valid extensions are either '{HSSF_EXTENSION}' or '{XSSF_EXTENSION}'.");
            }

            if (TypeUtil.IsPrimitive(typeof(T)))
            {
                PrimitiveTypeExport(book, exportCollections, excelOpt);
            }
            else
            {
                ClassTypeExport(book, exportCollections, excelOpt);
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
        /// <param name="book">excel sheet.</param>
        /// <param name="exportCollections">export any collections.</param>
        /// <param name="options">excel sheet options.</param>
        protected virtual void PrimitiveTypeExport<T>(IWorkbook book, IList<T> exportCollections, ExcelSpreadSheetOptions options)
        {
            var rowIndex = 0;
            var columnIndex = 0;

            if (options != null)
            {
                rowIndex = (int)options.StartRowIndex;
                columnIndex = (int)options.StartColumnIndex;
            }

            var sheet = book.CreateSheet(settings.SheetName);
            var style = book.CreateCellStyle();

            foreach (var value in exportCollections)
            {
                setCellValue(sheet: sheet, style: style, rowIndex: rowIndex, columnIndex: columnIndex, value);
                rowIndex++;
            }
        }

        /// <summary>
        /// Outputs class type.
        /// </summary>
        /// <typeparam name="T">any class types.</typeparam>
        /// <param name="book">excel sheet.</param>
        /// <param name="exportCollections">export any collections.</param>
        /// <param name="options">excel sheet options.</param>
        protected virtual void ClassTypeExport<T>(IWorkbook book, IList<T> exportCollections, ExcelSpreadSheetOptions options) 
        {
            var rowIndex = 0;
            var columnIndex = 0;
            var gType = typeof(T);

            if (options != null)
            {
                rowIndex = (int)options.StartRowIndex;
                columnIndex = (int)options.StartColumnIndex;
            }

            var sheet = book.CreateSheet(settings.SheetName);
            var style = book.CreateCellStyle();
            style.DataFormat = book.CreateDataFormat().GetFormat("MM/dd/yyyy HH:mm:ss");

            var hColIndex = columnIndex;
            if (ReadHeaderInfo.Property == settings.ReadHeaderInfo || ReadHeaderInfo.PropertyAndField == settings.ReadHeaderInfo)
            {
                foreach (var prop in gType.GetProperties())
                {
                    var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                    var cell = row.GetCell(hColIndex) ?? row.CreateCell(hColIndex);
                    cell.SetCellValue(prop.Name);
                    hColIndex++;
                }
            }
            if (ReadHeaderInfo.Field == settings.ReadHeaderInfo || ReadHeaderInfo.PropertyAndField == settings.ReadHeaderInfo)
            {
                foreach (var field in gType.GetFields())
                {
                    var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                    var cell = row.GetCell(hColIndex) ?? row.CreateCell(hColIndex);
                    cell.SetCellValue(field.Name);
                    hColIndex++;
                }
            }


            rowIndex = rowIndex + 1;
            foreach (var value in exportCollections)
            {
                var colIndex = columnIndex;
                if (ReadHeaderInfo.Property == settings.ReadHeaderInfo || ReadHeaderInfo.PropertyAndField == settings.ReadHeaderInfo)
                {
                    foreach (var prop in gType.GetProperties())
                    {
                        setCellValue(sheet: sheet, style: style, rowIndex: rowIndex, columnIndex: colIndex, prop.GetValue(value));
                        colIndex++;
                    }
                }
                if (ReadHeaderInfo.Field == settings.ReadHeaderInfo || ReadHeaderInfo.PropertyAndField == settings.ReadHeaderInfo)
                {
                    foreach (var field in gType.GetFields())
                    {
                        setCellValue(sheet: sheet, style: style, rowIndex: rowIndex, columnIndex: colIndex, field.GetValue(value));
                        colIndex++;
                    }
                }
                rowIndex++;
            }
        }

        /// <summary>
        /// Set the value in the cell.
        /// </summary>
        /// <typeparam name="T">any primitive type.</typeparam>
        /// <param name="sheet">sheet instance.</param>
        /// <param name="style">cell style.</param>
        /// <param name="rowIndex">row index.</param>
        /// <param name="columnIndex">column index.</param>
        /// <param name="value">set value.</param>
        private void setCellValue<T>(ISheet sheet, ICellStyle style, int rowIndex, int columnIndex, T value)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
            
            if (value is bool)
            {
                var tValue = value as bool?;
                cell.SetCellValue(tValue.Value);
            }

            if (value is int || value is float || value is double)
            {
                var tValue = double.Parse(value.ToString());
                cell.SetCellValue(tValue);
            }

            if (value is string)
            {
                cell.SetCellValue(value.ToString());
            }

            if (value is DateTime)
            {
                var tValue = value as DateTime?;
                cell.CellStyle = style;
                cell.SetCellValue(tValue.Value);
            }
        }
    }
}
