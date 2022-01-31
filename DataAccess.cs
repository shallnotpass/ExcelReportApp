using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace ExcelReportApp
{
    public class DataAccess
    {
        private readonly string ConnectionString;

        public DataAccess(string connectionString)
        {
            this.ConnectionString = connectionString;
        }

        public async Task< List<Product>> ReadProductsFromDatabase()
        {
            string queryText = @"SELECT Links.Izdel, Links.IzdelUp, Izdel.Name, Links.kol, Izdel.Price  FROM Links JOIN Izdel ON Links.Izdel = Izdel.Id;";
            List<Product> result = new List<Product>();
            using (SqlConnection connection = new SqlConnection(this.ConnectionString))
            {
                using (SqlCommand command = new SqlCommand(queryText, connection))
                {
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            object[] row = new object[5];
                            reader.GetValues(row);
                            if(row[1] != DBNull.Value)
                                result.Add(new Product { 
                                    Id = (long)row[0], 
                                    ParentId = (long)row[1], 
                                    Name = (string)row[2], 
                                    Number = (int)row[3], 
                                    Cost = (decimal)row[4] });
                            else
                                result.Add(new Product
                                {
                                    Id = (long)row[0],
                                    ParentId = null,
                                    Name = (string)row[2],
                                    Number = (int)row[3],
                                    Cost = (decimal)row[4]
                                });

                        }
                        return result;
                    }
                }
            }
        }

    }
}
