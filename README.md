# ExcelR
ExcelR is a simple C# library that help to import export helper using npoi

# Get work book

var workBook = ExportHelper.GetWorkbook();

# Get work sheet

var worksheet = ExportHelper.GetWorksheet("sheet1");

Create a simple model

var model = new List<T>(); //  where T is some model

export model to workseet

worksheet.Write(model);

Save workbook to disk 

workbook.Save(somefilepath);
