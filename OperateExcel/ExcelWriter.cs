using System;
using System.Data;
using System.IO;
using System.Drawing;
using OfficeOpenXml;


namespace Excel
{
    public class ExcelWriter
    {
        public static Boolean ExportDataSet(DataSet ds, string filePath)
        {
            ExcelPackage package = new ExcelPackage();
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                package = writeToSheet(package, ds.Tables[i]);
            }
            package.SaveAs(new FileInfo(filePath));
            return true;
        }

        private static ExcelPackage writeToSheet(ExcelPackage package, DataTable dt)
        {
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add(dt.TableName);
            int rows = dt.Rows.Count;
            int cols = dt.Columns.Count;
            for (int i = 1; i <= rows; i++)
            {
                DataRow dr = dt.Rows[i - 1];
                for (int j = 1; j <= cols; j++)
                {
                    //sheet.Cells[i, j].Style.Numberformat.Format = ((dynamic)dr[j - 1])["format"];
                    sheet.Cells[i, j].Value = ((dynamic)dr[j-1])["text"];
                    if(!string.IsNullOrWhiteSpace(((dynamic)dr[j - 1])["color"]))
                    {
                        Color color = System.Drawing.ColorTranslator.FromHtml("#" + ((dynamic)dr[j - 1])["color"]);
                        sheet.Cells[i, j].Style.Font.Color.SetColor(color);
                    }
                }  
            }
            
            ExcelRange r = sheet.Cells[1, 1, cols, rows];
            r.AutoFitColumns();
            
            return package;
        }
    }
}
