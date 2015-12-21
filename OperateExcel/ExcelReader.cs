using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using OfficeOpenXml;


namespace Excel
{
    public class ExcelReader
    {
        private static void Initialize(string filePath, out ExcelPackage package, out ExcelWorkbook workbook)
        {
            if (!File.Exists(filePath))
            {
                //return null;
            }
            if (! filePath.EndsWith(".xls")| !filePath.EndsWith(".xlsx") )
            {
                //return null;
            }
            package = new ExcelPackage(new FileInfo(filePath));
            workbook = package.Workbook;
        }


        public static DataSet getAllSheets(string filePath)
        {
            ExcelPackage package;
            ExcelWorkbook workbook;
            Initialize(filePath, out package, out workbook);
            int sheetsCount = package.Workbook.Worksheets.Count;
            DataSet sheetsSet = new DataSet();
            if (sheetsCount > 0)
            {
                try
                {
                    for (int i = 1; i <= sheetsCount; i++)
                    {
                        if (workbook.Worksheets[i].Dimension == null)       // empty sheet is not equal to null, but its dimension is.
                        {
                            //Console.WriteLine("sheet {0} is empty.", i);
                        }
                        else
                        {
                            DataTable sheetData = getSpecifiedSheet(workbook, i);
                            sheetsSet.Tables.Add(sheetData);
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }
            }
            return sheetsSet;
        }

        public static DataTable getSpecifiedSheet(string filePath, int sheetIndex)
        {
            if(!File.Exists(filePath))
            {
                return null;
            }

            ExcelPackage package;
            ExcelWorkbook workbook;
            Initialize(filePath, out package, out workbook);

            DataTable sheetData = new DataTable();

            try
            {
                getSpecifiedSheet(workbook, sheetIndex);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            return sheetData;
        }

        private static DataTable getSpecifiedSheet(ExcelWorkbook workbook, int sheetIndex)
        {
            ExcelWorksheet sheet = workbook.Worksheets[sheetIndex];
            DataTable sheetData = new DataTable(sheet.Name);
            int rowCount = sheet.Dimension.End.Row;
            int colCount = sheet.Dimension.End.Column;

            for (int i = 1; i <= colCount; i ++)
            {
                object cell =  sheet.Cells[i, 1].Value;
                string columnName = cell != null ? cell.ToString() : string.Empty;
                DataColumn column = new DataColumn();
                column.DataType = Type.GetType("System.Object");
                sheetData.Columns.Add(column);
            }

            try
            {
                for (int i = 1; i <= rowCount; i ++)
                {
                    
                    DataRow row = sheetData.NewRow();
                    for (int j = 1; j <= colCount; j ++)
                    {
                        string textColor = string.Empty;
                        ExcelRange cell = sheet.Cells[i, j];
                        object cellText = cell.Text;
                        if(cellText == null)
                        {
                            cellText = "--";
                        }
                        else
                        {
                            textColor = cell.Style.Font.Color.Rgb;
                            cellText = cellText.ToString();
                        }
                        string textFormat = cell.Style.Numberformat.Format;
                        Dictionary<string, object> box = new Dictionary<string, object> { { "text", cellText }, { "color", textColor}, { "format", textFormat } };
                        row[j - 1] = box;
                    }
                    sheetData.Rows.Add(row);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            return sheetData;
            
        }
    }
}
