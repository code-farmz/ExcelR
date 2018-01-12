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
    public  class ExportHelper
    {
        private  IWorkbook _workbook;

       


        public ExportHelper()
        {
            _workbook = GetWorkbook();
        }


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public  IWorkbook GetWorkbook(string defaultSheetName = "Sheet1")
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
        public  ISheet GetWorkSheet(string name = "Sheet1")
        {
            if (_workbook == null)
            _workbook = new XSSFWorkbook();
          return  _workbook.GetSheet(name) ?? _workbook.CreateSheet(name);
        }

        
        
    }
}
