using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel
{
    public class ExcelReader
    {
        private static void Initialize(string fileName, out Excel.Application xlApp, out Excel.Workbook xlWorkBook)
        {

        }




        public static DataSet ImportDataFromAllSheets(string fileName)
        {
            if (!File.Exists(fileName))
            {
                return null;
            }

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Initialize(fileName, out xlApp, out xlWorkBook);

            int workSheetNum = xlWorkBook.Worksheets.Count;
            int sheetCount=0;
            DataSet sheets = new DataSet();
            //Console.WriteLine("Sheet number: " + workSheetNum);
            try
            {
                for (sheetCount = 1; sheetCount <= workSheetNum; sheetCount++)
                {
                    //Console.WriteLine("Reading sheet{0}: {1}", sheetCount, xlWorkBook.Sheets[sheetCount].Name);
                    DataTable sheetData = ExtractDataFromSingleSheet(xlWorkBook, sheetCount);
                    sheets.Tables.Add(sheetData);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                Dispose(fileName, ref xlApp, ref xlWorkBook);
            }

            return sheets;
        }

        public static DataTable ImprotDataFromSingleSheet(string fileName, int sheetIndex)
        {
            if(!File.Exists(fileName))
            {
                return null;
            }

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Initialize(fileName, out xlApp, out xlWorkBook);

            DataTable sheetData = null;

            try
            {
                sheetData = ExtractDataFromSingleSheet(xlWorkBook, sheetIndex);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                Dispose(fileName, ref xlApp, ref xlWorkBook);
            }

            return sheetData;
        }

        private static DataTable ExtractDataFromSingleSheet(Excel.Workbook xlWorkBook, int sheetCount)
        {
            int rowCount = 0;
            int columnCount = 0;
            DataTable sheetData = new DataTable();

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetCount);
            Excel.Range range = xlWorkSheet.UsedRange;
            sheetData.TableName = xlWorkSheet.Name;

            for (int i = 0; i < range.Columns.Count; i++)
            {
                //sheetData.Columns.Add(new DataColumn()); 
                object cell = (range.Cells[1, i+1] as Excel.Range).Value2;
                string columnName = cell != null ? (range.Cells[1, i+1] as Excel.Range).Value2.ToString() : String.Empty;
                DataColumn column = new DataColumn();
                column.ColumnName = columnName;
                column.DataType = Type.GetType("System.Object");
                sheetData.Columns.Add(column);
            }

            try
            {
                for (rowCount = 1; rowCount <= range.Rows.Count; rowCount++)
                {
                    DataRow row = sheetData.NewRow();
                    for (columnCount = 1; columnCount <= range.Columns.Count; columnCount++)
                    {
                        object cellText = (range.Cells[rowCount, columnCount] as Excel.Range).Value;
                        double textColor = 0;
                        string textFormat = "G/通用格式";
                        double bgColor = 0;
                        if (cellText == null)
                        {
                            cellText = String.Empty;
                        }
                        else
                        {
                            cellText = (range.Cells[rowCount, columnCount] as Excel.Range).Value.ToString();
                            textColor = (range.Cells[rowCount, columnCount] as Excel.Range).Font.Color;
                            textFormat = (range.Cells[rowCount, columnCount] as Excel.Range).NumberFormatLocal;
                            bgColor = (range.Cells[rowCount, columnCount] as Excel.Range).Interior.Color;
                        }
                        Dictionary<string, object> cell = new Dictionary<string, object> { { "text", cellText }, { "color", textColor}, { "format", textFormat }, { "bgColor", bgColor} };
                        row[columnCount - 1] = cell;
                    }
                    sheetData.Rows.Add(row);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                ReleaseObject(xlWorkSheet);
            }

            return sheetData;
        }

        private static void Dispose(string fileName, ref Excel.Application xlApp, ref Excel.Workbook xlWorkBook)
        {
            xlWorkBook.Close(false, fileName, null);
            xlApp.Quit();

            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);
        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
