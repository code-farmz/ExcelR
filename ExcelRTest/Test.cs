using System;
using System.Collections.Generic;
using System.Runtime.Remoting.Channels;
using ExcelR;
using ExcelR.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace ExcelRTest
{
    [TestClass]
    public class Test
    {
        [TestMethod]
        public void ExcelExportTester()
        {
           var data = GetList();
            if(!Directory.Exists($"{Environment.CurrentDirectory}/docs"))
                Directory.CreateDirectory($"{Environment.CurrentDirectory}/docs");
           data.ToExcel("Sheet1",ExportHelper.Style.H3,ExportHelper.Color.Aqua).Save($"{Environment.CurrentDirectory}/docs/test_{DateTime.Now.Millisecond}.xlsx");
        }
        [TestMethod]
        public void ExcelImportTester()
        {
            var sheet =
                ImportHelper.GetWorkSheet(
                    @"C:\Users\Cena\Documents\Visual Studio 2015\Projects\ExcelR\ExcelRTest\docs\abc.xlsx");

           var data= sheet.Read<TestModel>();
        }

        [TestMethod]
        public void CsvExportTester()
        {
            var data = GetList();
            data.ToCsv(@"D:\abcTest.csv");

        }

        [TestMethod]
        public void CsvImportTester()
        {
            var data = CsvHelper.ReadFromFile<TestModel>(@"D:\abcTest.csv");
        }

        private List<TestModel> GetList()
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
    }
}
