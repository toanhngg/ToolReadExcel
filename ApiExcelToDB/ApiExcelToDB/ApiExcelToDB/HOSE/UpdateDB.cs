using Microsoft.Extensions.Configuration;
using System.Data.SqlClient;
using System.Data;
using System;

namespace ApiExcelToDB.HOSE
{
    public class UpdateDB
    {
    
            IConfiguration config = new ConfigurationBuilder()
               .AddJsonFile("appsettingshsx.json", optional: true, reloadOnChange: true)
               .Build();
            public void InsertDB(DataTable table, string name)
            {
                string connectionString = config["ConnectionStrings:SQLConnection"];
                using (SqlConnection sqlConnection = new SqlConnection(connectionString))
                {
                    sqlConnection.Open();
                    try
                    {
                        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConnection))
                        {
                            bulkCopy.DestinationTableName = name;
                            foreach (DataColumn col in table.Columns)
                            {
                                bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                            }

                            bulkCopy.WriteToServer(table);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }

                }
            }
        
    }
}
