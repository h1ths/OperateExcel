using System;
using System.Data;
using System.IO;
using Excel;

namespace OperateExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string file = @"test.xlsx";
            DataSet sheets;
            sheets = ExcelReader.getAllSheets(file);
            for(int i = 0; i < sheets.Tables.Count; i ++)
            {
                DataTable dt = sheets.Tables[i];
                Console.WriteLine("sheet {0}, named: {1}", i + 1, dt.TableName);
                for(int j = 0; j < dt.Rows.Count; j++)
                {
                    for(int k = 0; k < dt.Columns.Count; k++)
                    {
                        Console.Write(((dynamic)dt.Rows[j][k])["text"] + " |");
                    }
                }
            }
            string outfile = @"newxlsx.xlsx";
            if (File.Exists(outfile))
            {
                File.Delete(outfile);
            }
            ExcelWriter.ExportDataSet(sheets, outfile);
            Console.WriteLine();
            Console.WriteLine("write success!");
            Console.ReadKey();
        }
    }
}
