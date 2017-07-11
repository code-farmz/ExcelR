# ExcelR 
[![NuGet Version](https://img.shields.io/badge/nuget-v1.1.0-blue.svg)](https://www.nuget.org/packages/ExcelR/) 

This project  helps to create/read xlsx/csv files in an easy way.Best thing is you can directly export model to xlsx or import xlsx to model

# Why ExcelR
With the help of ExcelR you can
- Read or write file from/to disk or stream
- Write you list of objects to file 
- Read data from file to your model
- Control the color/font of column write to the xlsx file
- Set diffrent heading style
- Can read coustom property name with help of excelprop attribute
- And a lot more .....

# Dependancy and required references 
- NPOI v2.3.0
- NPOI.O0XML.dll
- NPOI.OpenXml4Net.dll
- NPOI.OpenXmlFormats.dll
- ICSharpCode.SharpZipLib.dll

# Examples:-
* Lets we have test class and some sample data as follow
  
   ```
   public class TestClass
    {
        [ExcelRProp(Name = "Custom String name",ColTextColor = "Red")] 
        // ExcelRProp attribute provide a number of custom option for this column
        public string StringProp { get; set; }
        public bool BoolProp { get; set; }
        public DateTime? DateTimeProp { get; set; }
        public int IntProp { get; set; }
    }
    var sampleData = new List<TestClass>
        {
         new TestClass {BoolProp = true, DateTimeProp = DateTime.Now, 
         StringProp = "jitender", IntProp = 5},
         new TestClass {BoolProp = false, StringProp = "jitende4r", 
         IntProp = 1}
         };
    ```
## Write and save data to xlsx file

#### Method1:-
```
sampleData.ToExcel().Save(filePath);
   ```

#### Method2:-


* Get worksheet and write data to sheet as follow
   ```
         var sheet = ExportHelper.GetWorkSheet();//you can pass custom sheet 
         name 
        //File data in the sheet
         sheet.Write(dataToWrite);
    ```
* Save sheet to stream or disk
   ```
    var stream = sheet.ToStream();
    sheet.Save(filePath);
   ```

## Write and save data to xlsx file
 ```
sampleData.ToCsv(filePath);
 ```
   
## Read data from xlsx file or stream
* Get worksheet from file or stream
   ```
   var workSheet=ImportHelper.GetWorkSheet(filePath)
                     Or
    var workSheet=ImportHelper.GetWorkSheet(stream)
   ```
 * Read data from sheet
   ```
   var data= sheet.Read<TestModel>();
   ```
   
## Manually creating xlsx from complex models

   ```
        var sheet = ExportHelper.GetWorkSheet();//you can pass custom sheet name 
         var rowNo=0;
         //create header row
         var headerRow = sheet.CreateRow(rowNo++,Style.H1);
         //Set header values
         headerRow.SetValue(0,"String property")
         //Create data rows and fill data
         foreach(var item in sampleData){
         var dataRow = sheet.CreateRow(rowNo++,Style.H1);
         dataRow.SetValue(0,item.StringProp);
         }
         
         //Save to  file
         sheet.Woorkbook.Save(filePath);
         //Output to stream
          sheet.Woorkbook.ToStream();
         
   ```
 # Need any help drop your queries:- cena666999@gmail.com