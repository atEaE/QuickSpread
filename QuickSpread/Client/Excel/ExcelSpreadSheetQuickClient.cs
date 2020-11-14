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
        public static IQuickClient Build(this ClientBuilder _, string value)
        {
            return new ExcelSpreadSheetQuickClient();
        }
    }

    public class ExcelSpreadSheetQuickClient : IQuickClient
    {
        /// <summary>
        /// Export to a specified spreadsheet.
        /// </summary>
        /// <typeparam name="T">POCO with no internal List, Array or Class</typeparam>
        /// <param name="exportCollections">export collections</param>
        /// <exception cref="ArgumentNullException"></exception>
        public void Export<T>(IList<T> exportCollections)
        {
            if (exportCollections == null)
                throw new ArgumentNullException(nameof(exportCollections));

            if (!exportCollections.Any())
                return;

            string filePath = @"C:\Users\EtaAoki\Desktop\新しいフォルダー\sample.xlsx";

            IWorkbook book;
            var extension = Path.GetExtension(filePath);
            if (extension == ".xls")
            {
                book = new HSSFWorkbook();
            }
            else if (extension == ".xlsx")
            {
                book = new XSSFWorkbook();
            }
            else
            {
                throw new ApplicationException("CreateNewBook: invalid extension");
            }

            var sheet = book.CreateSheet("sheet1");

            if (TypeUtil.IsPrimitive(typeof(T)))
            {
                PrimitiveTypeExport(sheet, exportCollections);
            }
            else
            {
                ClassTypeExport(sheet, exportCollections);
            }

            using (var fs = new FileStream(filePath, FileMode.Create))
            {
                book.Write(fs);
            }
        }

        protected virtual void PrimitiveTypeExport<T>(ISheet sheet, IList<T> exportCollections)
        {
            var rowIndex = 0;
            foreach(var value in exportCollections)
            {
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                var cell = row.GetCell(0) ?? row.CreateCell(0);
                cell.SetCellValue(value.ToString());
                rowIndex++;
            }
        }

        protected virtual void ClassTypeExport<T>(ISheet sheet, IList<T> exportCollections) 
        {
            var gType = typeof(T);
            var properties = gType.GetProperties();
            var props = gType.GetFields();

            var hColIndex = 0;
            foreach (var prop in gType.GetFields())
            {
                var row = sheet.GetRow(0) ?? sheet.CreateRow(0);
                var cell = row.GetCell(hColIndex) ?? row.CreateCell(hColIndex);
                cell.SetCellValue(prop.Name);
                hColIndex++;
            }

            var rowIndex = 1;
            foreach (var value in exportCollections)
            {
                var colIndex = 0;
                foreach(var prop in gType.GetFields())
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
