using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ExcelR.Attributes;
using NPOI.SS.UserModel;

namespace ExcelR.Extensions
{
    /// <summary>
    /// 
    /// </summary>
    public static class Excel
    {

        /// <summary>
        /// Save workbook to given filePath
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="path"></param>
        public static void Save(this IWorkbook workbook, string path)
        {
            using (var fileData = new FileStream(path, FileMode.Create))
            {
                workbook.Write(fileData);
            }
        }

        /// <summary>
        /// Save sheet to given filePath
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="path"></param>
        public static void Save(this ISheet sheet, string path)
        {
            using (var fileData = new FileStream(path, FileMode.Create))
            {
                sheet.Workbook.Write(fileData);
            }
        }
        /// <summary>
        /// Export sheet data to stream
        /// </summary>
        /// <param name="sheet"></param>
        public static Stream ToStream(this ISheet sheet)
        {
            using (var stream = new MemoryStream())
            {
                sheet.Workbook.Write(stream);
                return stream;
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
                workbook.Write(stream);
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
                workbook.Write(stream);
                return stream.ToArray();
            }
        }

        /// <summary>
        /// Export workbook data to byte array
        /// </summary>
        /// <param name="sheet"></param>
        public static byte[] ToByteArray(this ISheet sheet)
        {
            using (var stream = new MemoryStream())
            {
                sheet.Workbook.Write(stream);
                return stream.ToArray();
            }
        }

        /// <summary>
        /// Write given model data to sheet supported data types(int,bool,datetime,string,double,float)
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="data"></param>
        /// <param name="headerStyle">Header row style</param>
        /// <param name="headerColor">Header row color</param>
        /// <param name="textStyle">text style for whole sheet</param>
        /// <param name="textColor">Text color</param>
        /// <returns></returns>
        public static ISheet Write<TModel>(this ISheet sheet, IList<TModel> data, ExportHelper.Style headerStyle = ExportHelper.Style.H2, ExportHelper.Color headerColor = ExportHelper.Color.Black, ExportHelper.Style textStyle = ExportHelper.Style.Normal, ExportHelper.Color textColor = ExportHelper.Color.Black)
        {
            sheet.SetHeader(data, headerStyle, headerColor);
            sheet.FillData(data, textStyle, textColor);
            return sheet;
        }

        /// <summary>
        /// Write given model data to excel sheet supported data types(int,bool,datetime,string,double,float)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="sheetName">By default it will export to sheet1</param>
        /// <param name="headerStyle">Header row style</param>
        /// <param name="headerColor">Header row color</param>
        /// <param name="textStyle">text style for whole sheet</param>
        /// <param name="textColor">Text color</param>
        /// <returns></returns>
        public static IWorkbook ToExcel<TModel>(this IList<TModel> data,string sheetName, ExportHelper.Style headerStyle = ExportHelper.Style.H2, ExportHelper.Color headerColor = ExportHelper.Color.Black, ExportHelper.Style textStyle = ExportHelper.Style.Normal, ExportHelper.Color textColor = ExportHelper.Color.Black)
        {
            var sheet = ExportHelper.GetWorkSheet(sheetName);
            sheet.SetHeader(data, headerStyle, headerColor);
            sheet.FillData(data, textStyle, textColor);
            return sheet.Workbook;
        }

        #region private methods
        private static void SetHeader<TModel>(this ISheet sheet, IList<TModel> data, ExportHelper.Style headerStyle = ExportHelper.Style.H2, ExportHelper.Color color = ExportHelper.Color.Black)
        {
            foreach (var propertyInfo in data.GetType().GenericTypeArguments[0].GetProperties().Where(Include).Select((info, index) => new { info, index }))
            {
                var headColor = color;
                var propName = propertyInfo.info.Name;
                var attribute = propertyInfo.info.GetCustomAttributes(typeof(ExcelRProp), false).FirstOrDefault();
                var attrVal = attribute as ExcelRProp;
                if (!string.IsNullOrEmpty(attrVal?.Name))
                    propName = attrVal.Name;
                try
                {
                    if (attrVal?.HeadTextColor != null)
                        headColor = (ExportHelper.Color)Enum.Parse(typeof(ExportHelper.Color), attrVal.HeadTextColor, true);
                }
                catch (Exception)
                {

                    throw new Exception($"{attrVal?.HeadTextColor} is not a valid color");
                }
                var hearderRow = sheet.CreateRow(0, headerStyle);
                hearderRow.SetValue(propertyInfo.index, propName, headerStyle, headColor);
            }
        }

        private static bool Include(PropertyInfo propertyInfo)
        {
            var attribute = propertyInfo.GetCustomAttributes(typeof(ExcelRProp), false).FirstOrDefault();
            if (attribute == null)
                return true;
            
            var attrVal = attribute as ExcelRProp;
            return !attrVal?.SkipExport ?? true;
        }


        private static void FillData<TModel>(this ISheet sheet, IList<TModel> data, ExportHelper.Style textStyle = ExportHelper.Style.Normal, ExportHelper.Color textColor = ExportHelper.Color.Black)
        {

            foreach (var modelInfo in data.Select((model, index) => new { index, model }))
            {
                var row = sheet.CreateRow(modelInfo.index + 1, textStyle);
                foreach (var propertyInfo in modelInfo.model.GetType().GetProperties().Where(Include).Select((info, index) => new { info, index }))
                {
                    var attribute = propertyInfo.info.GetCustomAttributes(typeof(ExcelRProp), false).FirstOrDefault();
                    var attrVal = attribute as ExcelRProp;
                    try
                    {
                        if (attrVal?.ColTextColor != null)
                            textColor = (ExportHelper.Color)Enum.Parse(typeof(ExportHelper.Color), attrVal.ColTextColor, true);
                    }
                    catch (Exception)
                    {

                        throw new Exception($"{attrVal?.ColTextColor} is not a valid color");
                    }

                    var propType = propertyInfo.info.PropertyType;
                    var propVal = propertyInfo.info.GetValue(modelInfo.model);
                    row.SetValue(propertyInfo.index, propType, propVal, textStyle, textColor);
                    textColor = ExportHelper.Color.Black;
                }
            }
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
        private static void SetValue(this IRow row, int colNnum, Type propType, object propVal, ExportHelper.Style style = ExportHelper.Style.Normal, ExportHelper.Color textColor = ExportHelper.Color.Black)
        {
            var cell = row.GetCell(colNnum) ?? row.CreateCell(colNnum);
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
        #endregion
    }
}
