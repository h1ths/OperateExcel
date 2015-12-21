using System;
using System.Data;
using System.IO;
using Excel;
using System.Collections.Generic;

namespace OperateExcel
{
    class Program
    {

        public static void regulateData(DataTable dt, int width)
        {
            if (dt.Columns.Count > width)
            {
                for (int i = width; i < dt.Columns.Count; i++)
                {
                    dt.Columns.RemoveAt(i);
                }
            }
            else
            {
                width = dt.Columns.Count;
            }
            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < width; i++)
                {
                    if (!string.IsNullOrEmpty(((dynamic)dr[i])["color"]))
                    {
                        ((dynamic)dr[i])["color"] = Int32.Parse(((dynamic)dr[i])["color"].Substring(2), System.Globalization.NumberStyles.HexNumber);
                    }


                    string text = ((dynamic)dr[i])["text"];
                    string format = ((dynamic)dr[i])["format"];
                    if (text.IndexOf(".") != -1 && text.IndexOf(".") == text.LastIndexOf("."))
                    {

                        if (format.IndexOf("%") != -1)
                        {
                            ((dynamic)dr[i])["text"] = (Math.Round(double.Parse(text), 4, MidpointRounding.AwayFromZero) * 100).ToString() + "%";

                        }
                        else
                        {
                            ((dynamic)dr[i])["text"] = Math.Round(double.Parse(text), 2, MidpointRounding.AwayFromZero).ToString();
                        }
                    }


                    dr[i] = new Dictionary<string, object> { { "text", ((dynamic)dr[i])["text"] }, { "color", ((dynamic)dr[i])["color"] } };
                }
            }
        }

        static void Main(string[] args)
        {
            string file = @"test.xlsx";
            DataSet sheets;
            sheets = ExcelReader.getAllSheets(file);
            //regulateData(sheets.Tables[0], 119);
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
                    Console.WriteLine();
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
