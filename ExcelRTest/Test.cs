using System;
using System.Collections.Generic;
using System.Runtime.Remoting.Channels;
using ExcelR;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelRTest
{
    [TestClass]
    public class Test
    {
        [TestMethod]
        public void ExportTester()
        {
            var sheet = ExportHelper.GetSheet();

            var data = GetList();
            sheet.Write(data,ExportHelper.Style.H3,ExportHelper.Color.Aqua);
            var stream = sheet.Workbook.ToByteArray();
            sheet.Workbook.Save(@"C:\Users\Cena\Documents\Visual Studio 2015\Projects\ExcelR\ExcelRTest\docs\abc.xlsx");
        }
        [TestMethod]
        public void ImportTester()
        {
            var sheet =
                ImportHelper.GetSheet(
                    @"C:\Users\Cena\Documents\Visual Studio 2015\Projects\ExcelR\ExcelRTest\docs\abc.xlsx");

           var data= sheet.Read<TestModel>();
        }
        private List<TestModel> GetList()
        {
            var list = new List<TestModel>
            {
                new TestModel {Bool = true, DateTime = DateTime.Now, String = "jitender", Int = 5},
                new TestModel {Bool = false, String = "jitende4r", Int = 1}
            };
            return list;
        }
    }
}
