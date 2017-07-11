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
     public class TestModel
    {
        [ExcelRProp(Name = "First Name")]
        public string FirstName { get; set; }


        [ExcelRProp(ColTextColor = "Red", Name = "Last Name")]
        public string LastName { get; set; }

        [ExcelRProp(SkipExport = true)]
        public bool IsMale { get; set; }

        [ExcelRProp(HeadTextColor = "Blue" ,Name = "Date Of Birth")]
        public DateTime? Dob { get; set; }

    }
    var sampleData = new List<TestModel>
            {
                new TestModel {IsMale = true, Dob = DateTime.Now, FirstName = "jitender", LastName = "kundu"},
                new TestModel {IsMale = true,  FirstName = "raj"},
                new TestModel {IsMale = true, Dob = DateTime.Now.AddDays(15), FirstName = "Michel"},
                new TestModel {IsMale = true, Dob = DateTime.Now, FirstName = "Cena", LastName = "raj"},
                new TestModel {IsMale = false, FirstName = "john", LastName = "Cena"}
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

## Write and save data to csv file
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
## Read data from csv file
```
 var data = CsvHelper.ReadFromFile<TestModel>(sourceFilePath);
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