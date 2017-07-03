# ExcelR
This project helps to create/read xlsx files in an easy way.

# Why ExcelR
With the help of ExcelR you can
- Read or write file from/to disk or stream
- Write you list of objects to file 
- Read data from file to your model
- Control the color/font of column write to the xlsx file
- Set diffrent heading style
- Can read coustom property name with help of excelprop attribute
- and a lot more.....

# Examples:-
* Create a test class
   ```
   public class TestClass
    {
        [ExcelRProp(Name = "Custom String name",ColTextColor = "Red")] 
        // you may skip above attribute
        public string StringProp { get; set; }
        public bool BoolProp { get; set; }
        public DateTime? DateTimeProp { get; set; }
        public int IntProp { get; set; }
    }
    ```
## Write and save data to xlsx file


* Get worksheet and write data to sheet as follow
   ```
         var sheet = ExportHelper.GetSheet();//you can pass custom sheet 
         name 
         var dataToWrite = new List<TestModel>
        {
         new TestModel {BoolProp = true, DateTimeProp = DateTime.Now, 
         StringProp = "jitender", IntProp = 5},
         new TestModel {BoolProp = false, StringProp = "jitende4r", 
         IntProp = 1}
         };
         sheet.Write(dataToWrite);
    ```
* Save sheet to stream or disk
   ```
    var stream = sheet.Workbook.ToStream();
    sheet.Workbook.Save(filePath);
   ```
   
## Read data from xlsx file or stream
* Get worksheet from file or stream
   ```
   var workSheet=ImportHelper.GetSheet(filePath)
                     Or
    var workSheet=ImportHelper.GetSheet(stream)
   ```
 * Read data from sheet
   ```
   var data= sheet.Read<TestModel>();
   ```
 Happy coding...............................