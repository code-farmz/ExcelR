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
            /// <summary>
            /// 
            /// </summary>
            Normal,
            /// <summary>
            /// Bold font with text size 9
            /// </summary>
            Bold,
            /// <summary>
            /// Bold font with text size 22
            /// </summary>
            H1,
            /// <summary>
            /// Bold font with text size 16
            /// </summary>
            H2,
            /// <summary>
            /// Bold font with text size 10
            /// </summary>
            H3
        }
        #endregion



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
