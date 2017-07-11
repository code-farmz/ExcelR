using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ExcelR.Attributes;
using NPOI.HSSF.Record.Chart;
using NPOI.SS.UserModel;

namespace ExcelR.Extensions
{
    /// <summary>
    /// Contains extension methods to create csv
    /// </summary>
    public static class CsvHelper
    {
        private static IDictionary<string, int> _propColumnMap;
        #region Write methods

        /// <summary>
        /// export give list of items to byte array
        /// </summary>
        /// <param name="data"></param>
        /// <param name="delimiter">Separator to used default will be ,</param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static byte[] ToCsv<T>(this IList<T> data, char delimiter = ',')
        {
            using (var memoryStream = new MemoryStream())
            using (var writer = new StreamWriter(memoryStream))
            {
                writer.SetHeader(data, delimiter);
                writer.FillData(data, delimiter);
                writer.Flush();
                return memoryStream.ToArray();
            }
        }
        /// <summary>
        /// export give list of items to output stream
        /// </summary>
        /// <param name="data"></param>
        /// <param name="delimiter">Separator to used default will be ,</param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static Stream ToCsvStream<T>(this IList<T> data, char delimiter = ',')
        {
            using (var memoryStream = new MemoryStream())
            using (var writer = new StreamWriter(memoryStream))
            {
                writer.SetHeader(data, delimiter);
                writer.FillData(data, delimiter);
                writer.Flush();
                return memoryStream;
            }
        }

        /// <summary>
        /// export give list of items to given file path
        /// </summary>
        /// <param name="data"></param>
        /// <param name="filePath"></param>
        /// <param name="delimiter">Separator to used default will be ,</param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static void ToCsv<T>(this IList<T> data, string filePath, char delimiter = ',')
        {
            using (var writer = new StreamWriter(filePath))
            {
                writer.SetHeader(data, delimiter);
                writer.FillData(data, delimiter);
                writer.Flush();
            }
        }

        /// <summary>
        /// Save stream to give file path
        /// </summary>
        /// <param name="bytesArray"></param>
        /// <param name="filePath"></param>
        public static void Save(this byte[] bytesArray, string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new Exception("File path is required");
            using (var stream = new MemoryStream(bytesArray))
            using (var fileStream = File.Create(filePath))
            {
                stream.Seek(0, SeekOrigin.Begin);
                stream.CopyTo(fileStream);

            }

        }

        private static void SetHeader<TModel>(this StreamWriter writer, IList<TModel> data, char delimiter)
        {
            string row = null;
            foreach (var propertyInfo in data.GetType().GenericTypeArguments[0].GetProperties().Where(Include).Select((info, index) => new { info, index }))
            {
                var propName = propertyInfo.info.Name;
                var attribute = propertyInfo.info.GetCustomAttributes(typeof(ExcelRProp), false).FirstOrDefault();
                var attrVal = attribute as ExcelRProp;
                if (!string.IsNullOrEmpty(attrVal?.Name))
                    propName = attrVal.Name;

                row = row != null ? $"{row}{delimiter}{propName}" : propName;
            }
            writer.WriteLine(row);
        }

        private static void FillData<TModel>(this StreamWriter writer, IList<TModel> data, char delimiter)
        {
            foreach (var row in data.Select((model, index) => new { index, model }).Select(modelInfo => modelInfo.model.GetType().GetProperties().Where(Include).Select((info, index) => new { info, index }).Select(propertyInfo => propertyInfo.info.GetValue(modelInfo.model)).Aggregate<object, string>(null, (current, propVal) => current != null ? $"{current}{delimiter}{propVal}" : propVal?.ToString())))
            {
                writer.WriteLine(row);
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

        #endregion

        #region ReadMethods

        /// <summary>
        /// work in progress
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath">Source file path</param>
        /// <param name="delimiter"></param>
        /// <returns></returns>
        public static IList<T> ReadFromFile<T>(string filePath, char delimiter = ',')
        {
           using (var streamReader = new StreamReader(filePath))
            {
                var retVal = new List<T>();
                Activator.CreateInstance<T>().SetPropColumnMapings(streamReader, delimiter);
                if (_propColumnMap == null || !_propColumnMap.Any())
                    return retVal;
                string row;
                var rowno=0;
                while ((row = streamReader.ReadLine()) != null)
                {
                    rowno += 1;
                    if (rowno <= 1) continue;
                    var keys = row.Split(delimiter);
                    var model = Activator.CreateInstance<T>();
                    foreach (var propertyInfo in model.GetType().GetProperties())
                    {
                        var propType = propertyInfo.PropertyType;
                        var name = propertyInfo.Name;
                        var attribute = propertyInfo.GetCustomAttributes(typeof(ExcelRProp), false).FirstOrDefault();
                        var attrVal = attribute as ExcelRProp;
                        if (!string.IsNullOrEmpty(attrVal?.Name))
                            name = attrVal.Name;
                        if (!_propColumnMap.ContainsKey(name)) continue;
                        var index = _propColumnMap[name];
                        var matchingValue = keys[index];
                        if(string.IsNullOrEmpty(matchingValue)) continue;
                        if (propType == typeof(string))
                        {
                            propertyInfo.SetValue(model, matchingValue);
                        }
                        else if (propType == typeof(bool) || propType == typeof(bool?))
                        {
                            bool val;
                            if(bool.TryParse(matchingValue, out val))
                            propertyInfo.SetValue(model, val);
                        }
                        else if (propType == typeof(int) || propType == typeof(int?))
                        {
                            int val;
                            if (int.TryParse(matchingValue, out val))
                                propertyInfo.SetValue(model, val);
                        }
                        else if (propType == typeof(double) || propType == typeof(double?))
                        {
                            double val;
                            if (double.TryParse(matchingValue, out val))
                                propertyInfo.SetValue(model, val);
                        }
                        else if (propType == typeof(float) || propType == typeof(float?))
                        {
                            float val;
                            if (float.TryParse(matchingValue, out val))
                                propertyInfo.SetValue(model, val);
                        }
                        else if (propType == typeof(DateTime) || propType == typeof(DateTime?))
                        {
                            DateTime val;
                            if (DateTime.TryParse(matchingValue, out val))
                                propertyInfo.SetValue(model, val);
                        }
                    }
                    retVal.Add(model);
                }
                return retVal;
            }
        }

        private static void SetPropColumnMapings<TModel>(this TModel model, StreamReader streamReader, char delimiter)
        {
            _propColumnMap = new Dictionary<string, int>();
            var keys = new string[] { };
            string row;
            while ((row = streamReader.ReadLine()) != null)
            {
                keys = row.Split(delimiter);
                break;
            }

            if (!keys.Any()) return;
            foreach (var propertyInfo in model.GetType().GetProperties())
            {
                var name = propertyInfo.Name;
                var attribute = propertyInfo.GetCustomAttributes(typeof(ExcelRProp), false).FirstOrDefault();
                var attrVal = attribute as ExcelRProp;
                if (!string.IsNullOrEmpty(attrVal?.Name))
                    name = attrVal.Name;
                if (!keys.Any(key => key.ToLower().Equals(name.ToLower()))) continue;
                var index = Array.IndexOf(keys, name);
                _propColumnMap.Add(name, index);
            }
        }

        #endregion
    }
}
