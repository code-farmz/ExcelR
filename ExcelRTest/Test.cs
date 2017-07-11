using System;
using System.Collections.Generic;
using System.Runtime.Remoting.Channels;
using ExcelR;
using ExcelR.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelRTest
{
    [TestClass]
    public class Test
    {
        [TestMethod]
        public void ExcelExportTester()
        {
           var data = GetList();
           data.ToExcel("Sheet1",ExportHelper.Style.H3,ExportHelper.Color.Aqua).Save(@"C:\Users\Cena\Documents\Visual Studio 2015\Projects\ExcelR\ExcelRTest\docs\abc.xlsx");
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

        private List<TestModel> GetList()
        {
            var list = new List<TestModel>
            {
                new TestModel {Bool = true, DateTime = DateTime.Now, String = "jitender", Int = 5},
                new TestModel {Bool = true,  String = "raj"},
                new TestModel {Bool = true, DateTime = DateTime.Now.AddDays(15), String = "jit", Int = 45},
                new TestModel {Bool = true, DateTime = DateTime.Now, String = "cena", Int = 455},
                new TestModel {Bool = false, String = "john", Int = 1}
            };
            return list;
        }
    }
}
