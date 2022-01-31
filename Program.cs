// See https://aka.ms/new-console-template for more information
using ExcelReportApp;
DataAccess dataAccess = new DataAccess("Data Source=DESKTOP-5T6AFIE;Initial Catalog=Products;Integrated Security=True");
ExcelAccess excelAccess = new ExcelAccess();
await excelAccess.WriteReport(await dataAccess.ReadProductsFromDatabase());
