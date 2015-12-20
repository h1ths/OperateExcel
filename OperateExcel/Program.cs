using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using System.Data;

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
                for(int j = 0; j < dt.Rows.Count; j++)
                {
                    for(int k = 0; k < dt.Columns.Count; k++)
                    {
                        Console.Write(dt.Rows[i][k] + " ");
                    }
                    Console.WriteLine();
                }
            }
            Console.ReadKey();
        }
    }
}
