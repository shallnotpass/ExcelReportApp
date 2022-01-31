using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Tables;

namespace ExcelReportApp
{
    public class ExcelAccess
    {
        public async Task<string> WriteReport(List<Product> products)
        {
            List<Product> orderedProducts = OrderProducts(products);
            List<Product> outputProducts = new List<Product>();
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            var workbook = new ExcelFile();
            var worksheet = workbook.Worksheets.Add("Tables");
            var data = new object[products.Count, 3];
            foreach (Product product in orderedProducts)
            {
                outputProducts.Add(product);
                if (product.Children != null)
                {
                    foreach (Product partProduct in product.Children)
                    {
                        outputProducts.Add(partProduct);
                        if (partProduct.Children != null)
                            outputProducts.AddRange(partProduct.Children);
                    }
                }
            }
            for (int i = 0; i < products.Count; i++)
            {
                worksheet.Cells[i, 0].Value = outputProducts[i].Name;
                worksheet.Cells[i, 1].Value = outputProducts[i].Number;
                worksheet.Cells[i, 2].Value = outputProducts[i].Cost;
            }
            worksheet.Columns[0].SetWidth(100, LengthUnit.Pixel);
            worksheet.Columns[1].SetWidth(70, LengthUnit.Pixel);
            worksheet.Columns[2].SetWidth(70, LengthUnit.Pixel);
            worksheet.Columns[3].SetWidth(70, LengthUnit.Pixel);
            worksheet.Columns[2].Style.NumberFormat = "#,##0.00";
            worksheet.Columns[3].Style.NumberFormat = "#,##0.00";
            try
            {
                workbook.Save("Report.xlsx");
            }
            catch (Exception ex)
            {
                return "Fail";
            }
            return "Success";
        }

        private List<Product> OrderProducts(List<Product> products)
        {
            List<Product> orderedProducts = products.Where(x=>x.ParentId==null).ToList();
            foreach (Product product in orderedProducts)
            {
                product.Children = products.Where(x => x.ParentId == product.Id).ToList();
                foreach(Product partProduct in product.Children)
                {
                    partProduct.Children = products.Where(x => x.ParentId == partProduct.Id).ToList();
                    partProduct.Children.ForEach(x => x.Name = String.Concat("        ", x.Name));
                    partProduct.Cost = partProduct.Number * (partProduct.Cost + partProduct.Children.Sum(x => x.Cost));
                    partProduct.Name = String.Concat("        ", partProduct.Name);
                }
                product.Cost = product.Number * (product.Cost + product.Children.Sum(x => x.Cost));
            }
            return orderedProducts;
        }
    }
}
