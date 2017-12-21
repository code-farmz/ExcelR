using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ExcelR.ExportHelper;

namespace ExcelR.Extensions
{
  public static  class ExcelExport
    {
        private static IDictionary<IFont, ICellStyle> _styles;


        static ExcelExport()
        {
            _styles = new Dictionary<IFont, ICellStyle>();
        }

        /// <summary>
        /// Get the sheet with given name will search in given workbook will create if exist
        /// </summary>
        /// <returns></returns>
        public static ISheet GetWorkSheet(this IWorkbook workbook, string sheetName = "Sheet1")
        {
            return workbook.GetSheet(sheetName) ?? workbook.CreateSheet(sheetName);
        }

        /// <summary>
        /// Add new row to given sheet
        /// </summary>
        /// <param name="sheet">ISheet</param>
        /// <param name="rownum">Row Number</param>
        /// <param name="style">Enum</param>
        /// <param name="color">Text color</param>
        /// <returns></returns>
        public static IRow CreateRow(this ISheet sheet, int rownum, Style style = Style.Normal, Color color = Color.Black)
        {
            var row = sheet.GetRow(rownum) ?? sheet.CreateRow(rownum);
            row.SetStyle(style);
            if (color != Color.Black)
            {
                var font = row.RowStyle.GetFont(sheet.Workbook);
                font.Color = IndexedColors.ValueOf(color.ToString()).Index;
                row.RowStyle.SetFont(font);
            }
            return row;
        }

        /// <summary>
        /// Set value in given cell no
        /// </summary>
        /// <param name="row"></param>
        /// <param name="colNnum"></param>
        /// <param name="value"></param>
        /// <param name="style"></param>
        /// <param name="color">Text color to be set</param>
        public static void SetValue(this IRow row, int colNnum, string value, Style style = Style.Normal, Color color = Color.Black)
        {
            var cell = row.CreateCell(colNnum);
            cell.SetStyle(style, color);
            cell.SetCellValue(value);
        }

        /// <summary>
        /// Set value in given cell no
        /// </summary>
        /// <param name="row"></param>
        /// <param name="colNnum"></param>
        /// <param name="value"></param>
        /// <param name="style"></param>
        /// <param name="color">Text color to be set</param>
        public static void SetValue(this IRow row, int colNnum, bool value, Style style = Style.Normal, Color color = Color.Black)
        {
            var cell = row.CreateCell(colNnum);
            cell.SetStyle(style, color);
            cell.SetCellValue(value);
        }

        /// <summary>
        /// Set date value in given cell no
        /// </summary>
        /// <param name="row"></param>
        /// <param name="colNnum"></param>
        /// <param name="value"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        public static void SetValue(this IRow row, int colNnum, DateTime value, Style style = Style.Normal, Color color = Color.Black)
        {
            var cell = row.CreateCell(colNnum);
            cell.SetStyle(style, color);
            cell.SetCellValue(value);
        }


        /// <summary>
        /// Set number value in given cell no
        /// </summary>
        /// <param name="row"></param>
        /// <param name="colNnum"></param>
        /// <param name="value"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        public static void SetValue(this IRow row, int colNnum, double value, Style style = Style.Normal, Color color = Color.Black)
        {
            var cell = row.CreateCell(colNnum);
            cell.SetStyle(style, color);
            cell.SetCellValue(value);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="style"></param>
        /// <param name="textColor"></param>
        public static void SetStyle(this ICell cell, Style style, Color textColor = Color.Black)
        {
            cell.CellStyle =cell.Sheet.Workbook.GetStyle(style, textColor);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="style"></param>
        public static void SetStyle(this IRow row, Style style)
        {
            row.RowStyle = row.Sheet.Workbook.GetStyle(style);
        }
        /// <summary>
        /// Add new sheet to the workbook
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static ISheet CreateSheet(this IWorkbook workbook, string sheetName)
        {
            return workbook.CreateSheet(sheetName);
        }

        #region privatemethods
        private static ICellStyle GetStyle(this IWorkbook workbook, Style style, Color color = Color.Black)
        {
            switch (style)
            {
                case Style.H1:
                    return workbook.GetStyle(22, color, FontBoldWeight.Bold);
                case Style.H2:
                    return workbook.GetStyle(16, color, FontBoldWeight.Bold);
                case Style.H3:
                    return workbook.GetStyle(11, color, FontBoldWeight.Bold);
                default:
                    return workbook.GetStyle(10, color);
            }
        }

        private static ICellStyle GetStyle(this IWorkbook workbook,short fontSize = 9, Color color = Color.Black, FontBoldWeight fontBoldWeight = FontBoldWeight.None, string fontName = "Calibri")
        {
            var font = workbook.GetFont(fontSize, fontName, color, fontBoldWeight);
            return workbook.GetStyle(font);
        }

        private static ICellStyle GetStyle(this IWorkbook workbook ,IFont font)
        {
            if (_styles.Any(s => s.Key.Equals(font)))
                return _styles[font];
            var style = workbook.CreateCellStyle();
            style.SetFont(font);
            return style;
        }
        private static IFont GetFont(this IWorkbook workbook,short fontSize, string fontName, Color color = Color.Black, FontBoldWeight fontBoldWeight = FontBoldWeight.None)
        {
            var font = workbook.CreateFont();
            font.FontHeightInPoints = fontSize;
            font.FontName = fontName;
            font.Boldweight = (short)fontBoldWeight;
            font.Color = IndexedColors.ValueOf(color.ToString()).Index;
            return font;
        }



        #endregion

    }
}
