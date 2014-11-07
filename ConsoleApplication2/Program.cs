using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            dt.Clear();
            dt.Columns.Add("Coluna 1");
            dt.Columns.Add("Coluna 2");
            dt.Columns.Add("Coluna 3");
            dt.Columns.Add("Coluna 4");
            dt.Columns.Add("VALIDACAO DATA");
            dt.Columns.Add("Lista");
            DataRow dr = dt.NewRow();
            dr["Coluna 1"] = "Valor 1";
            dr["Coluna 2"] = "Valor 2";
            dr["Coluna 3"] = "Valor 3";
            dr["Coluna 4"] = "Valor 4";
            dr["VALIDACAO DATA"] = "1986-03-26";
            dr["Lista"] = "Selecione algume item"; 
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["Coluna 1"] = "Valor Linha 2 - 1";
            dr["Coluna 2"] = "Valor Linha 2 - 2";
            dr["Coluna 3"] = "Valor Linha 2 - 3";
            dr["Coluna 4"] = "ESSA PODE SER ALTERADA";
            dr["VALIDACAO DATA"] = "DATA";

            dt.Rows.Add(dr);

            DumpExcel(dt);
        }

        public static void DumpExcel(DataTable tbl)
        {
            using (ExcelPackage pck = new ExcelPackage())
            {
                //Create the worksheet
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("TabelaX");

                //Load the datatable into the sheet, starting from cell A1. Print the column names on row 1
                ws.Cells["A1"].LoadFromDataTable(tbl, true);

                //Format the header for column 1-3
                using (ExcelRange rng = ws.Cells["A1:B1"])
                {
                    rng.Style.Font.Bold = true;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));  //Set color to dark blue
                    rng.Style.Font.Color.SetColor(Color.White);
                }

                using (ExcelRange rng = ws.Cells["D1"])
                {
                    rng.Style.Font.Bold = true;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));  //Set color to dark blue
                    rng.Style.Font.Color.SetColor(Color.White);
                }

                using (ExcelRange rng = ws.Cells["A2"])
                {
                    rng.Style.Font.Bold = true;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                    rng.Style.Fill.BackgroundColor.SetColor(Color.Red);  //Set color to dark blue
                    rng.Style.Font.Color.SetColor(Color.White);
                }

                ws.Cells["E2"].Style.Locked = false;
                ws.Cells["D3"].Style.Locked = false;
                ws.Cells["E3"].Style.Locked = false;
                ws.Cells["F2"].Style.Locked = false;
                ws.Protection.SetPassword("123");
                ws.Protection.IsProtected = true;
                ws.Cells[ws.Dimension.Address].AutoFitColumns();

                //Example how to Format Column 1 as numeric 
                using (ExcelRange col = ws.Cells[2, 1, 2 + tbl.Rows.Count, 1])
                {
                    col.Style.Numberformat.Format = "#,##0.00";
                    col.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }

                var val = ws.DataValidations.AddListValidation("F2:F10");
                val.Formula.Values.Add("1 - Test do Id 1");
                val.Formula.Values.Add("2 - Test do Id 2");
                val.Formula.Values.Add("3 - Test do Id 3");
                val.Formula.Values.Add("4 - Test do Id 4");
                val.Formula.Values.Add("5 - Test do Id 5");
                val.Formula.Values.Add("6 - Test do Id 6");
                val.Formula.Values.Add("7 - Test do Id 7");

                // Add a date time validation
                ValidateDate(ws, "E2");
                ValidateDate(ws, "E3");

                ws.Column(1).Hidden = true;
                //Write it back to the client
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("content-disposition", "attachment;  filename=ExcelDemo.xlsx");
                //Response.BinaryWrite(pck.GetAsByteArray());
                try
                {
                    FileStream fs = File.Create("C:\\teste.xlsx", 2048, FileOptions.None);
                    BinaryWriter bw = new BinaryWriter(fs);

                    byte[] ba = pck.GetAsByteArray();

                    bw.Write(ba);

                    bw.Close();
                    fs.Close();
                }
                catch (Exception e)
                {
                    Console.Write(e.Message);
                    Console.ReadKey(true);
                }

            }
        }

        private static void ValidateDate(ExcelWorksheet ws, string sheet)
        {
            var validation = ws.DataValidations.AddDateTimeValidation(sheet);
            // set validation properties
            validation.ShowErrorMessage = true;
            validation.ErrorTitle = "An invalid date was entered";
            validation.Error = "The date must be between 2011-01-31 and 2011-12-31";
            validation.Prompt = "Enter date here";
            validation.Formula.Value = DateTime.Parse("2011-01-01");
            validation.Formula2.Value = DateTime.Parse("2011-12-31");
            validation.Operator = ExcelDataValidationOperator.between;
        }
    }
}
