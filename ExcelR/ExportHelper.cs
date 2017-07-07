using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using ExcelR.Attributes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelR
{
    public static class ExportHelper
    {
        private static IWorkbook _workbook;

        private static IDictionary<IFont, ICellStyle> _styles; 
        #region Enums
        public enum Color
        {
            Aqua,
            Automatic,
            Black,
            Blue,
            BlueGrey,
            BrightGreen,
            Brown,
            Coral,
            CornflowerBlue,
            DarkBlue,
            DarkGreen,
            DarkRed,
            DarkTeal,
            DarkYellow,
            Gold,
            Green,
            Grey25Percent,
            Grey40Percent,
            Grey50Percent,
            Grey80Percent,
            Indigo,
            Lavender,
            LemonChiffon,
            LightBlue,
            LightCornflowerBlue,
            LightGreen,
            LightOrange,
            LightTurquoise,
            LightYellow,
            Lime,
            Maroon,
            OliveGreen,
            Orange,
            Orchid,
            PaleBlue,
            Pink,
            Plum,
            Red,
            Rose,
            RoyalBlue,
            SeaGreen,
            SkyBlue,
            Tan,
            Teal,
            Turquoise,
            Violet,
            White,
            Yellow

        }

        public enum Style
        {
            /// <summary>
            /// 
            /// </summary>
            Normal,
            /// <summary>
            /// Bold font with text size 9
            /// </summary>
            Bold,
            /// <summary>
            /// Bold font with text size 22
            /// </summary>
            H1,
            /// <summary>
            /// Bold font with text size 16
            /// </summary>
            H2,
            /// <summary>
            /// Bold font with text size 10
            /// </summary>
            H3
        }
        #endregion



        static ExportHelper()
        {
            _workbook = GetWorkbook();
            _styles = new Dictionary<IFont, ICellStyle>();
        }


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static IWorkbook GetWorkbook(string defaultSheetName = "Sheet1")
        {
            if (_workbook != null)
                return _workbook;
            _workbook = new XSSFWorkbook();
            _workbook.CreateSheet(defaultSheetName);
            return _workbook;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static ISheet GetWorkSheet(string name = "Sheet1")
        {
            if (_workbook == null)
            _workbook = new XSSFWorkbook();
          return  _workbook.GetSheet(name) ?? _workbook.CreateSheet(name);
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

      

        /// <summary>
        /// Get the sheet with given name will search in given workbook will create if exist
        /// </summary>
        /// <returns></returns>
        public static ISheet GetWorkSheet(this IWorkbook workbook,string sheetName = "Sheet1")
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
                var font = row.RowStyle.GetFont(_workbook);
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



        #region private methods

      

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="style"></param>
        /// <param name="textColor"></param>
        public static void SetStyle(this ICell cell, Style style, Color textColor = Color.Black)
        {
            cell.CellStyle = GetStyle(style,textColor);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="style"></param>
        public static void SetStyle(this IRow row, Style style)
        {
            row.RowStyle = GetStyle(style);
        }

        private static ICellStyle GetStyle(Style style,Color color = Color.Black)
        {
            switch (style)
            {
                case Style.H1:
                   return GetStyle(22, color, FontBoldWeight.Bold);
                case Style.H2:
                    return GetStyle(16, color, FontBoldWeight.Bold);
                case Style.H3:
                    return GetStyle(11, color, FontBoldWeight.Bold);
                default:
                    return GetStyle(10, color);
            }
        }

        private static ICellStyle GetStyle(short fontSize=9, Color color = Color.Black, FontBoldWeight fontBoldWeight= FontBoldWeight.None, string fontName = "Calibri")
        {
            var font = GetFont(fontSize, fontName, color, fontBoldWeight);
            return GetStyle(font);
        }

        #endregion

        #region private props

        private static ICellStyle GetStyle(IFont font)
        {
            if (_styles.Any(s => s.Key.Equals(font)))
                return _styles[font];
            var style = _workbook.CreateCellStyle();
            style.SetFont(font);
            return style;
        }
        private static IFont GetFont(short fontSize, string fontName, Color color = Color.Black, FontBoldWeight fontBoldWeight = FontBoldWeight.None)
        {
            var font = _workbook.CreateFont();
            font.FontHeightInPoints = fontSize;
            font.FontName = fontName;
            font.Boldweight = (short)fontBoldWeight;
            font.Color = IndexedColors.ValueOf(color.ToString()).Index;
            return font;
        }


        #endregion

    }
}
