using NoExcelImportTools;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var excelTools = new ExcelImport();
            string filePath = args[0];
            var data = excelTools.ExcelToDataTable(filePath);
            for (int i = 0; i < data.Rows.Count; i++)
            {
                DataRow row = data.Rows[i];
                for (int c = 0; c < row.Table.Columns.Count; c++)
                {
                    Console.Write(row[c].ToString());
                    Console.Write(",");
                }
                Console.WriteLine();
            }
            Console.Read();
        }
    }
}
