﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelR.Attributes;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelR
{
    public static class ImportHelper
    {

        private static IDictionary<string, string> _propColumnMap;
        public static ISheet GetSheet(this IWorkbook workbook, string sheetName= "Sheet1")
        {
            return workbook.GetSheet(sheetName);
        }

        public static ISheet GetSheet(Stream stream, string sheetName="Sheet1")
        {
            var workbook = GetWorkbook(stream);
            return workbook.GetSheet(sheetName);
        }

        public static ISheet GetSheet(string filePath, string sheetName = "Sheet1")
        {
            var workbook = GetWorkbook(filePath);
            return workbook.GetSheet(sheetName);
        }
        public static IWorkbook GetWorkbook(Stream stream)
        {
            return new XSSFWorkbook(stream);
        }

        public static IWorkbook GetWorkbook(string filePath)
        {
            if(!File.Exists(filePath))
                throw new Exception("File not found");
            using (var stream= new MemoryStream(File.ReadAllBytes(filePath)))
            {
                return new XSSFWorkbook(stream);
            }
        }

        /// <summary>
        /// Read data from given sheet and fill to given TModel
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static IList<TModel> Read<TModel>(this ISheet sheet)
        {
            var list = new List<TModel>();

            if (sheet == null)
                throw new Exception("Sheet must not be null");
            var headerRow = sheet.GetRow(0);
            if (headerRow == null)
            {
                throw new Exception("No row found at position 0");
            }
            var rows = sheet.GetRowEnumerator();
            var mapDict = Activator.CreateInstance<TModel>().GetSetPropColumnMapings(headerRow);
            while (rows.MoveNext())
            {
                IRow row;
                try
                {
                    row = (XSSFRow)rows.Current;
                }
                catch (Exception)
                {
                    row = (HSSFRow)rows.Current;
                }
                if (row.RowNum == 0) continue;
                var model = Activator.CreateInstance<TModel>();
                foreach (var propertyInfo in model.GetType().GetProperties())
                {
                    var propType = propertyInfo.PropertyType;
                    var name = propertyInfo.Name;
                    var attribute = propertyInfo.GetCustomAttributes(typeof(ExcelRProp), false).FirstOrDefault();
                    var attrVal = attribute as ExcelRProp;
                    if (!string.IsNullOrEmpty(attrVal?.Name))
                        name = attrVal.Name;
                    var cellNo = mapDict[name];
                    if (cellNo != null)
                    {
                        var cell = row.GetCell(int.Parse(cellNo), MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                        {
                            if (propType == typeof(string))
                            {
                                propertyInfo.SetValue(model, GetStringCellValue(cell));
                            }
                            else if (propType == typeof(bool) || propType == typeof(bool?))
                            {
                                propertyInfo.SetValue(model, GetBooleanCellValue(cell));
                            }
                            else if (propType == typeof(int) || propType == typeof(int?))
                            {
                                propertyInfo.SetValue(model, Convert.ToInt32(GetStringCellValue(cell)));
                            }
                            else if (propType == typeof(double) || propType == typeof(double?))
                            {
                                propertyInfo.SetValue(model, Convert.ToDouble(GetStringCellValue(cell)));
                            }
                            else if (propType == typeof(float) || propType == typeof(float?))
                            {
                                var val = GetStringCellValue(cell);
                                if (!string.IsNullOrEmpty(val))
                                    propertyInfo.SetValue(model, float.Parse(val));
                            }
                            else if (propType == typeof(DateTime) || propType == typeof(DateTime?))
                            {
                                propertyInfo.SetValue(model, GetDateTimeCellValue(cell));
                            }
                        }
                    }
                }
                list.Add(model);
            }
            return list;
        }

        private static IDictionary<string, string> GetSetPropColumnMapings<TModel>(this TModel model, IRow headerRow)
        {
            if (_propColumnMap != null && _propColumnMap.Count > 0)
                return _propColumnMap;
            _propColumnMap = new Dictionary<string, string>();
            foreach (var propertyInfo in model.GetType().GetProperties())
            {
                var name = propertyInfo.Name;
                var attribute = propertyInfo.GetCustomAttributes(typeof(ExcelRProp), false).FirstOrDefault();
                if (attribute != null)
                {
                    var attrVal = attribute as ExcelRProp;
                    if (!string.IsNullOrEmpty(attrVal?.Name))
                        name = attrVal.Name;
                }
                var matchingCell =
                    headerRow.Cells.FirstOrDefault(
                        cell =>
                            cell.CellType == CellType.String &&
                            cell.StringCellValue.ToLower().Equals(name.ToLower()));
                if (matchingCell != null)
                {
                    _propColumnMap.Add(name, matchingCell.ColumnIndex.ToString());
                }
            }

            return _propColumnMap;
        }


        public static string GetStringCellValue(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Numeric:
                    return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);

                default:
                    return null;
            }
        }

        public static bool? GetBooleanCellValue(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                default:
                    return null;
            }
        }

        public static DateTime? GetDateTimeCellValue(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.String:
                    return DateTime.Parse(cell.StringCellValue);

                case CellType.Numeric:
                        return cell.DateCellValue;

                default:
                    return null;
            }
        }


    }
}
