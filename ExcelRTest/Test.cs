using System;
using System.Collections.Generic;
using ExcelR;
using ExcelR.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using static ExcelR.Enums.Excel;

namespace ExcelRTest
{
    [TestClass]
    public class Test
    {
        [TestMethod]
        public string ExportExcel()
        {
            var filePath = $"{DocsDirctory}/test_{DateTime.Now.Millisecond}.xlsx";
            var data = GetSampleData();
            data.ToExcel("Sheet1", Style.H3, Color.Aqua).Save(filePath);
            return filePath;
        }
        [TestMethod]
        public void ImportExcel()
        {
           var sheet =ImportHelper.GetWorkSheet(ExportExcel());
           var data= sheet.Read<TestModel>();
        }


        [TestMethod]
        public string ExportCsv()
        {
            var filePath = $"{DocsDirctory}/test_{DateTime.Now.Millisecond}.csv";
            var data = GetSampleData();
            data.ToCsv(filePath);
            return filePath;

        }

        [TestMethod]
        public void ImportCsv()
        {
            var data = CsvHelper.ReadFromFile<TestModel>(ExportCsv());
        }

        private List<TestModel> GetSampleData()
        {
            var list = new List<TestModel>
            {
                new TestModel {IsMale = true, Dob = DateTime.Now, FirstName = "jitender", LastName = "kundu"},
                new TestModel {IsMale = true,  FirstName = "raj"},
                new TestModel {IsMale = true, Dob = DateTime.Now.AddDays(15), FirstName = "Michel"},
                new TestModel {IsMale = true, Dob = DateTime.Now, FirstName = "Cena", LastName = "raj"},
                new TestModel {IsMale = false, FirstName = "john", LastName = "Cena"}
            };
            return list;
        }

        #region private methods/props
        private string DocsDirctory
        {
            get
            {
                if (!Directory.Exists($"{Environment.CurrentDirectory}/docs"))
                    Directory.CreateDirectory($"{Environment.CurrentDirectory}/docs");
                return $"{Environment.CurrentDirectory}/docs";
            }
        }
        
        #endregion
    }
}
