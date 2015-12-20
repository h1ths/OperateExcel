using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using NPOI;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace Excel
{
    public class ExcelReader
    {
        private static void Initialize(string filePath, out IWorkbook workbook)
        {
            if (!File.Exists(filePath))
            {
                //return null;
            }
            if (! filePath.EndsWith(".xls")| !filePath.EndsWith(".xlsx") )
            {
                //return null;
            }
            using (FileStream fs = File.OpenRead(filePath))
            {
                workbook = WorkbookFactory.Create(fs);
            }
        }


        public static DataSet getAllSheets(string filePath)
        {
            IWorkbook workbook;
            Initialize(filePath, out workbook);
            int sheetsCount = workbook.NumberOfSheets;
            DataSet sheetsSet = new DataSet();
            Console.WriteLine("Sheet number: " + sheetsCount);
            try
            {
                for (int i  = 0; i < sheetsCount; i++)
                {
                    DataTable sheetData = getOneSheet(workbook, i);
                    sheetsSet.Tables.Add(sheetData);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return sheetsSet;
        }

        public static DataTable getOneSheet(string filePath, int sheetIndex)
        {
            if(!File.Exists(filePath))
            {
                return null;
            }

            IWorkbook workbook;
            Initialize(filePath, out workbook);

            DataTable sheetData = new DataTable();

            try
            {
                getOneSheet(workbook, sheetIndex);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            return sheetData;
        }

        private static DataTable getOneSheet(IWorkbook workbook, int sheetIndex)
        {
            DataTable sheetData = new DataTable();
            ISheet sheet = workbook.GetSheetAt(sheetIndex);
            int rowCount = sheet.LastRowNum;
            int columnCount = sheet.GetRow(0).LastCellNum;
            for (int i = 0; i < rowCount; i++)
            {
                if(columnCount < sheet.GetRow(i).LastCellNum)
                {
                    columnCount = sheet.GetRow(i).LastCellNum;
                }
            }

            for (int i = 0; i < columnCount; i ++)
            {
                object cell = sheet.GetRow(0).GetCell(i);
                string columnName = cell != null ? cell.ToString() : string.Empty;
                DataColumn column = new DataColumn();
                column.ColumnName = columnName;
                column.DataType = Type.GetType("System.Object");
                sheetData.Columns.Add(column);
            }

            try
            {
                for (int i = 0; i < rowCount; rowCount ++)
                {
                    DataRow row = sheetData.NewRow();
                    for (int j = 0; j < columnCount; j++)
                    {
                        object cellText = sheet.GetRow(i).GetCell(j);
                        Console.WriteLine(cellText);
                        double textColor = 0;
                        string textFormat = "G/通用格式";
                        double bgColor = 0;
                        if (cellText == null)
                        {
                            cellText = String.Empty;
                        }
                        else
                        {
                            cellText = sheet.GetRow(i).GetCell(j).ToString();
                            /*
                            textColor = sheet.GetRow(rowIndex).GetCell(columnIndex).CellStyle.GetFont();
                            textFormat = 
                            bgColor = 
                            */
                        }
                        Dictionary<string, object> cell = new Dictionary<string, object> { { "text", cellText }, { "color", textColor}, { "format", textFormat }, { "bgColor", bgColor} };
                        row[j] = cellText;
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
