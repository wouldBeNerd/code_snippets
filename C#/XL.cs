using OfficeOpenXml;
using System;
using System.Data;
using System.IO;
using System.Linq;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;

namespace SQL_to_GRAPH_v2_2021
{
    static class XL
    {
        public static ExcelPackage get_file(string filename)
        {
            FileInfo fi = new FileInfo(filename);
            ExcelPackage excelPackage = new ExcelPackage(fi);
            return excelPackage;
        }
        public static void worksheet_delete(ExcelPackage excelPackage, string sheetname)
        {
            try
            {
                excelPackage.Workbook.Worksheets.Delete(sheetname);
            }
            catch (Exception)
            {
                Console.WriteLine($"no sheet to delete, {sheetname}");
                //throw;
            }
        }
        public static ExcelWorksheet worksheet_add(ExcelPackage excelPackage, string sheetname)
        {
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(sheetname);
            return worksheet;
        }
        public static void DT_to_worksheet(ExcelWorksheet ws, DataTable dt, string table_name)
        {

            List<string> XL_columnnames = dt.Columns.Cast<DataColumn>()
                     .Select(x => x.ColumnName)
                     .ToList();

            int row_i = 1;
            int col_i = 1;
            foreach (string col in XL_columnnames)
            {
                ws.Cells[row_i, col_i].Value = col;
                col_i = col_i + 1;

            }
            row_i = row_i + 1;
                foreach (DataRow row in dt.Rows)
            {
                col_i = 1;
                foreach( string col in XL_columnnames)
                {

                    ws.Cells[row_i, col_i].Value = row[col];

                    col_i = col_i + 1;
                }
                col_i = 1;
                row_i = row_i + 1;
            }

            //no longer making a table from the range

            //ExcelRange range = ws.Cells[1, 1, row_i - 1, XL_columnnames.Count];
            //ExcelTable tab = ws.Tables.Add(range, table_name);
            //tab.TableStyle = TableStyles.Dark1;


        }
        public static void DA_to_XL(string filename, string sheetname, DataTable DT )
        {
            FileInfo fi = new FileInfo(filename);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                excelPackage.Workbook.Worksheets.Delete(sheetname);
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(sheetname);

                worksheet.Cells[1, 1].Value = "col A";
                worksheet.Cells[1, 2].Value = "col B";
                worksheet.Cells[1, 3].Value = "col C";

                worksheet.Cells[2, 1].Value = "red";
                worksheet.Cells[2, 2].Value = true;
                worksheet.Cells[2, 3].Value = 15;

                worksheet.Cells[3, 1].Value = "yellow";
                worksheet.Cells[3, 2].Value = false;
                worksheet.Cells[3, 3].Value = 12;

                worksheet.Cells[4, 1].Value = "turquoise";
                worksheet.Cells[4, 2].Value = true;
                worksheet.Cells[4, 3].Value = 3;

                worksheet.Cells[5, 1].Value = "green";
                worksheet.Cells[5, 2].Value = false;
                worksheet.Cells[5, 3].Value = 20;

                worksheet.Cells[6, 1].Value = "blue";
                worksheet.Cells[6, 2].Value = true;
                worksheet.Cells[6, 3].Value = 20;

                worksheet.Cells[7, 1].Value = "black";
                worksheet.Cells[7, 2].Value = true;
                worksheet.Cells[7, 3].Value = 20;


                worksheet.Cells[8, 1].Value = "white";
                worksheet.Cells[8, 2].Value = true;
                worksheet.Cells[8, 3].Value = 12;


                worksheet.Cells[9, 1].Value = "grey";
                worksheet.Cells[9, 2].Value = true;
                worksheet.Cells[9, 3].Value = 7;


                ExcelRange range = worksheet.Cells[1, 1, 9, 3];
                ExcelTable tab = worksheet.Tables.Add(range, "Table1");
                tab.TableStyle = TableStyles.Dark1;

                //Save your file
                excelPackage.SaveAs(fi);
            }
        }


        public static void create()
        {
                    //Create a new ExcelPackage
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
               //Set some properties of the Excel document
               excelPackage.Workbook.Properties.Author = "VDWWD";
               excelPackage.Workbook.Properties.Title = "Title of Document";
               excelPackage.Workbook.Properties.Subject = "EPPlus demo export data";
               excelPackage.Workbook.Properties.Created = DateTime.Now;

                //Create the WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

                //Add some text to cell A1
                worksheet.Cells["A1"].Value = "My first EPPlus spreadsheet!";
                //You could also use [line, column] notation:
                worksheet.Cells[1, 2].Value = "This is cell B1!";

                //Save your file
                FileInfo fi = new FileInfo(@"File.xlsx");
                excelPackage.SaveAs(fi);
            }

        }
        public static void open()
        {

            //Opening an existing Excel file
            FileInfo fi = new FileInfo(@"File.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                //Get a WorkSheet by index. Note that EPPlus indexes are base 1, not base 0!
                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[1];

                //Get a WorkSheet by name. If the worksheet doesn't exist, throw an exeption
                ExcelWorksheet namedWorksheet = excelPackage.Workbook.Worksheets["SomeWorksheet"];

                //If you don't know if a worksheet exists, you could use LINQ,
                //So it doesn't throw an exception, but return null in case it doesn't find it
                ExcelWorksheet anotherWorksheet =
                    excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == "SomeWorksheet");

                //Get the content from cells A1 and B1 as string, in two different notations
                string valA1 = firstWorksheet.Cells["A1"].Value.ToString();
                string valB1 = firstWorksheet.Cells[1, 2].Value.ToString();

                //Save your file
                excelPackage.Save();
            }
        }





    }
}
