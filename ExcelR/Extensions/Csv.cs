using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ExcelR.Attributes;
using NPOI.HSSF.Record.Chart;

namespace ExcelR.Extensions
{
    /// <summary>
    /// Contains extension methods to create csv
    /// </summary>
    public static class Csv
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
        public static void ToCsv<T>(this IList<T> data, string filePath,char delimiter = ',')
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
        /// <param name="filePath"></param>
        /// <param name="delimiter"></param>
        /// <returns></returns>
        private static IList<T> ReadFromCsv<T>(string filePath,char delimiter=',')
        {
            using (var streamReader = new StreamReader(filePath))
            {
                string row;
                while ((row = streamReader.ReadLine()) != null)
                {
                    
                }
            }
            return new List<T>();
        }

        private static void SetMappings(string headerRow)
        {
            
        }

        #endregion
    }
}
