using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bytescout.Spreadsheet;

namespace UploadCSV
{
    static class Program
    {
        #region IMPORT CSV FILE TO THE DATABASE
        static void Main(string[] args)
        {
            var connectionString = ConfigurationManager.ConnectionStrings["csvFileContext"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    #region DATABASE QUERY
                    //// Drop test database if exists
                    //ExecuteQueryWithoutResult(connection, "IF DB_ID('CSV_FILE_TEST') IS NOT NULL DROP DATABASE CSV_FILE");
                    //// Create empty database
                    //ExecuteQueryWithoutResult(connection, "CREATE DATABASE CSV_FILE_TEST");

                    // Switch to created database
                    //ExecuteQueryWithoutResult(connection, "USE CSVtoDB"); // change "CSVtoDB" to database name

                    //// Create a table for CSV data
                    //ExecuteQueryWithoutResult(connection,
                    //"CREATE TABLE [dbo].[cr_expences](iCtr INT,cCRCode CHAR(50),cExpenses CHAR(50),iCost DECIMAL(9, 2))");
                    #endregion

                    using (Spreadsheet document = new Spreadsheet())
                    {
                        //Console.WriteLine("Please wait...");

                        document.LoadFromFile(ConfigurationManager.AppSettings["path"], ","); //path of the file

                        Worksheet worksheet = document.Workbook.Worksheets[0];

                        try
                        {
                            for (int row = 0; row <= worksheet.UsedRangeRowMax; row++)
                            {
                                try
                                {
                                    String insertCommand = string.Format(
                                        "INSERT INTO crCashReceipt VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}',{7},{8},{9},{10},'{11}','{12}','{13}','{14}','{15}', {16},'{17}')",
                                        worksheet.Cell(row, 0).Value,
                                        worksheet.Cell(row, 1).Value,
                                        worksheet.Cell(row, 2).ValueAsExcelDisplays,
                                        worksheet.Cell(row, 3).Value,
                                        worksheet.Cell(row, 4).Value,
                                        worksheet.Cell(row, 5).Value,
                                        worksheet.Cell(row, 6).Value,
                                        worksheet.Cell(row, 7).Value,
                                        worksheet.Cell(row, 8).Value,
                                        worksheet.Cell(row, 9).Value,
                                        worksheet.Cell(row, 10).Value,
                                        worksheet.Cell(row, 11).ValueAsExcelDisplays,
                                        worksheet.Cell(row, 12).ValueAsExcelDisplays,
                                        worksheet.Cell(row, 13).Value,
                                        worksheet.Cell(row, 14).Value,
                                        worksheet.Cell(row, 15).Value,
                                        worksheet.Cell(row, 16).Value,
                                        worksheet.Cell(row, 17).Value);

                                    ExecuteQueryWithoutResult(connection, insertCommand);

                                }
                                catch
                                {
                                    String insertCommand = string.Format(
                                        "INSERT INTO crCashReceipt VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}',{7},{8},{9},{10},'{11}','{12}','{13}','{14}','{15}', {16},'{17}')",
                                        worksheet.Cell(row, 0).Value,
                                        worksheet.Cell(row, 1).Value,
                                        worksheet.Cell(row, 2).ValueAsExcelDisplays,
                                        worksheet.Cell(row, 3).Value,
                                        worksheet.Cell(row, 4).Value,
                                        worksheet.Cell(row, 5).Value,
                                        worksheet.Cell(row, 6).Value,
                                        worksheet.Cell(row, 7).Value,
                                        worksheet.Cell(row, 8).Value,
                                        worksheet.Cell(row, 9).Value,
                                        worksheet.Cell(row, 10).Value,
                                        worksheet.Cell(row, 11).ValueAsExcelDisplays,
                                        worksheet.Cell(row, 12).ValueAsExcelDisplays,
                                        worksheet.Cell(row, 13).Value,
                                        worksheet.Cell(row, 14).Value,
                                        worksheet.Cell(row, 15).Value,
                                        0,                            // because cFreight is null or empty
                                        worksheet.Cell(row, 17).Value);

                                    ExecuteQueryWithoutResult(connection, insertCommand);
                                }

                                Console.WriteLine();
                                Console.WriteLine($"Uploaded data: {row + 1}");
                            }
                        }

                        //catch (DbException exe)
                        //{
                        //    Console.WriteLine("Error: " + exe.Message);
                        //    Console.ReadKey();
                        //}

                        catch (Exception ex)
                        {
                            Console.WriteLine("Error: " + ex.Message);
                            Console.ReadKey();
                        }

                        Console.WriteLine();
                        Console.WriteLine("Successfully uploaded");
                        Console.ReadKey();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                    Console.ReadKey();
                }
            }

        }
        static void ExecuteQueryWithoutResult(SqlConnection connection, string query)
        {
            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.ExecuteNonQuery();
                //command.ExecuteNonQueryAsync();
            }
        }
        #endregion
    }

}
