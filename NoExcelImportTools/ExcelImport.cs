using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NoExcelImportTools
{
    public class ExcelImport
    {
        public DataTable ExcelToDataTable(string filePath)
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            IWorkbook workbook = application.Workbooks.Open(filePath);
            IWorksheet sheet = workbook.Worksheets[0];
            DataTable dataTable = sheet.ExportDataTable(sheet.UsedRange, ExcelExportDataTableOptions.None);

            return dataTable;
        }
    }
}
