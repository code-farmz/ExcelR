using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelR;
using ExcelR.Attributes;

namespace ExcelRTest
{
    public class TestModel
    {
        [ExcelRProp(Name = "TestString")]
        public string String { get; set; }

        public bool Bool { get; set; }

        public DateTime? DateTime { get; set; }

        [ExcelRProp(ColTextColor = "Red")]
        public int Int { get; set; }
    }

}
