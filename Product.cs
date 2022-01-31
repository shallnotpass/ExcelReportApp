using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportApp
{
    public class Product
    {
        public long? Id { get; set; }
        public long? ParentId { get; set; }
        public string? Name { get; set; }
        public int? Number { get; set; }
        public decimal? Cost { get; set; }
        public List<Product>? Children { get; set;}
    }
}
