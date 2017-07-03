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
        /// Custom color for the column text for all rows except header row while exporting available colors (Aqua,Automatic,Black,Blue,BlueGrey,BrightGreen,Brown,Coral,CornflowerBlue,DarkBlue,DarkGreen,DarkRed,DarkTeal,DarkYellow,
        /// Gold,Green,Grey25Percent,Grey40Percent,Grey50Percent,Grey80Percent,Indigo,Lavender,LemonChiffon,LightBlue,LightCornflowerBlue,LightGreen,LightOrange,LightTurquoise,LightYellow,Lime,
        /// Maroon,OliveGreen,Orange,Orchid,PaleBlue,Pink,Plum,Red,Rose,RoyalBlue,SeaGreen,SkyBlue,Tan,Teal,Turquoise,Violet,White,Yellow)
        ///  </summary>
       public string ColTextColor { get; set; }
        

    }
}
