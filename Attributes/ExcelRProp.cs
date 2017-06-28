using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelR.Attributes
{
  /// <summary>
  /// 
  /// </summary>
  public  class ExcelRProp : Attribute
    {
      /// <summary>
      /// Custom name of the property that will be used while importing or exporting
      /// </summary>
      public string Name { get; set; }

        /// <summary>
        /// Custom color for the column text for all rows except header row while exporting
        /// </summary>
      public ExportHelper.Color TextColor { get; set; }
    }
}
