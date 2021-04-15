using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.IO;
using System.Threading.Tasks;

namespace SQL_to_GRAPH_v2_2021
{


    public class SQL

    {
        public SqlConnection Connection()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = "server.domainname.com";
            builder.UserID = "username";
            builder.Password = "userpassword";
            builder.InitialCatalog = "DATABASE_name";
            builder.ConnectTimeout = 60000;
            return new SqlConnection(builder.ConnectionString);
        }
        public static void close(SqlConnection connection){
            connection.Close();
        }
        /// <summary>
        /// THIS IS BY NO MEANS A BEST PRACTICE, BEYOND RUNNING A SCHEDULED PROGRAM, I WOULD NOT RECOMMEND T HIS METHOD
        /// </summary>
        /// <param name="connection"></param>
        /// <param name="SQL_query"></param>
        /// <returns></returns>
        public async Task<DataSet> Run(SqlConnection connection, string SQL_query)
        {


                try
                {
                    using (connection)
                    {
                        Console.WriteLine("\nQuery data example:");
                        Console.WriteLine("=========================================\n");

                        SqlDataAdapter adapter = new SqlDataAdapter();
                        DataSet dataset = new DataSet();
                        using (adapter.SelectCommand = new SqlCommand(SQL_query, connection))
                        {
                            await connection.OpenAsync();
                            adapter.SelectCommand.CommandTimeout = 60000;
                            Console.WriteLine("running fill");
                            await Task.Run(() => 
                                adapter.Fill(dataset) 
                            );
                            connection.Close();
                            return dataset;

                        }
                    }
                }
                catch (SqlException e)
                {
                    Console.WriteLine(e.ToString());
                    DataSet dataset = new DataSet();
                    return dataset;
                }

        }
        /// <summary>
        /// import and read a sql file
        /// </summary>
        /// <param name="path">"starts at the destination where your Program.cs file is located"</param>
        /// <returns></returns>
        public string ImportFile(string path)
        {
  
            string sql_q = "";
            if (System.IO.File.Exists(path))
            {
                sql_q = System.IO.File.ReadAllText(path);
            }
            else
            {
                throw new System.ArgumentException($"{path} is not valid, SQL query file not found");
            }

            return sql_q;
        }

    }

}

