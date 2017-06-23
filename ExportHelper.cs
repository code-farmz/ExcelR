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
        private static ICellStyle _boldStyle;
        private static ICellStyle _h1Style;
        private static ICellStyle _h2Style;
        private static ICellStyle _h3Style;

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
            Normal,
            Bold,
            H1,
            H2,
            H3
        }
        #endregion



        static ExportHelper()
        {
            _workbook = GetWorkbook();
            _boldStyle = BoldStyle;
            _h1Style = H1Style;
            _h2Style = H2Style;
            _h3Style = H3Style;
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
        /// Write given model data to sheet supported data types(int,bool,datetime,string,double,float)
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="data"></param>
        /// <param name="headerStyle">Header row style</param>
        /// <param name="textStyle">text style</param>
        /// <returns></returns>
        public static ISheet Write<TModel>(this ISheet sheet, IList<TModel> data, Style headerStyle = Style.H2, Color headerColor = Color.Black,Style textStyle = Style.Normal, Color textColor = Color.Black)
        {
            sheet.SetHeader(data, headerStyle,headerColor);
            sheet.FillData(data,textStyle,textColor);
            return sheet;
        }

        private static void SetHeader<TModel>(this ISheet sheet, IList<TModel> data, Style headerStyle = Style.H2,Color color=Color.Black)
        {
            foreach (var propertyInfo in data.GetType().GenericTypeArguments[0].GetProperties().Select((info, index) => new { info, index }))
            {
                var propName = propertyInfo.info.Name;
                var attribute = propertyInfo.info.GetCustomAttributes(typeof(ExcelRProp), false).FirstOrDefault();
                var attrVal = attribute as ExcelRProp;
                if (!string.IsNullOrEmpty(attrVal?.Name))
                    propName = attrVal.Name;
                var hearderRow = sheet.CreateRow(0, headerStyle);
                hearderRow.SetValue(propertyInfo.index, propName, headerStyle, color);
            }
        }

        private static void FillData<TModel>(this ISheet sheet, IList<TModel> data, Style textStyle = Style.Normal,  Color textColor = Color.Black)
        {
            
            foreach (var modelInfo in data.Select((model, index) => new { index, model }))
            {
                var row = sheet.CreateRow(modelInfo.index + 1, textStyle);
                foreach (var propertyInfo in modelInfo.model.GetType().GetProperties().Select((info, index) => new { info, index }))
                {
                    var attribute = propertyInfo.info.GetCustomAttributes(typeof(ExcelRProp), false).FirstOrDefault();
                    var attrVal = attribute as ExcelRProp;
                    try
                    {
                        if (attrVal?.ColTextColor != null)
                            textColor = (Color)Enum.Parse(typeof(Color), attrVal.ColTextColor, true);
                    }
                    catch (Exception)
                    {
                        
                        throw new Exception($"{attrVal.ColTextColor} is not a valid color");
                    }
                   
                    var propType = propertyInfo.info.PropertyType;
                    var propVal = propertyInfo.info.GetValue(modelInfo.model);
                    row.SetValue(propertyInfo.index, propType,propVal,textStyle,textColor);
                    textColor = Color.Black;
                }
            }
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
        /// Save workbook to given filePath
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="path"></param>
        public static void Save(this IWorkbook workbook, string path)
        {
            using (var fileData = new FileStream(path, FileMode.Create))
            {
                _workbook.Write(fileData);
            }
        }

        /// <summary>
        /// Export workbook data to stream
        /// </summary>
        /// <param name="workbook"></param>
        public static Stream ToStream(this IWorkbook workbook)
        {
            using (var stream = new MemoryStream())
            {
                _workbook.Write(stream);
                return stream;
            }
        }

        /// <summary>
        /// Export workbook data to byte array
        /// </summary>
        /// <param name="workbook"></param>
        public static byte[] ToByteArray(this IWorkbook workbook)
        {
            using (var stream = new MemoryStream())
            {
                _workbook.Write(stream);
                return stream.ToArray();
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static ISheet GetSheet(string sheetName = "Sheet1")
        {
            return GetWorkbook().GetSheet(sheetName) ?? _workbook.CreateSheet(sheetName);
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
        public static void SetValue(this IRow row, int colNnum, double value, Style style = Style.Normal, Color color = Color.Black)
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
        /// <param name="propVal"></param>
        /// <param name="propType"></param>
        /// <param name="style"></param>
        /// <param name="textColor"></param>
        private static void SetValue(this IRow row, int colNnum, Type propType, object propVal, Style style = Style.Normal, Color textColor = Color.Black)
        {
            var cell = row.GetCell(colNnum)?? row.CreateCell(colNnum);
            cell.SetStyle(style, textColor);
            
            if (propVal == null)
            {
                return;
            }
                if (propType == typeof(string))
            {
                cell.SetCellValue(propVal.ToString());
            }
            else if (propType == typeof(bool))
            {
                cell.SetCellValue(Convert.ToBoolean(propVal.ToString()));
            }
            else if (propType == typeof(int))
            {
                cell.SetCellValue(Convert.ToInt32(propVal.ToString()));
            }
            else if (propType == typeof(double) || propType == typeof(float))
            {
                cell.SetCellValue(Convert.ToDouble(propVal.ToString()));
            }
            else if (propType == typeof(DateTime) || propType == typeof(DateTime?))
            {
                cell.SetCellValue(Convert.ToDateTime(propVal.ToString()));
            }
            
        }

        #region private methods
        private static void SetStyle(this ICell cell, Style style, Color textColor = Color.Black)
        {
            cell.CellStyle = GetStyle(style,textColor);
        }
        private static void SetStyle(this IRow row, Style style)
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
                    return GetStyle(12, color, FontBoldWeight.Bold);
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
        private static ICellStyle BoldStyle
        {
            get
            {
                var font = GetFont(11, "Calibri", Color.Black, FontBoldWeight.Bold);
                return GetStyle(font);
            }

        }

        private static ICellStyle H1Style
        {
            get
            {
                var font = GetFont(22, "Calibri", Color.Black, FontBoldWeight.Bold);
                return GetStyle(font);
            }

        }
        private static ICellStyle H2Style
        {
            get
            {
                var font = GetFont(16, "Calibri", Color.Black, FontBoldWeight.Bold);
                return GetStyle(font);
            }

        }

        private static ICellStyle H3Style
        {
            get
            {
                var font = GetFont(10, "Calibri", Color.Black, FontBoldWeight.Bold);
                return GetStyle(font);
            }

        }

        private static ICellStyle GetStyle(IFont font)
        {
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
