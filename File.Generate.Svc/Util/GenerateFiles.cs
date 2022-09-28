using ClosedXML.Excel;
using Project.Generate.Svc.Models;
using System.Data;
using System.Text;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;

namespace Project.Generate.Svc.Util
{
    public static class GenerateFiles
    {
        public static DataTable GenerateDataTable(IEnumerable<Client> clientes)
        {
            var table = new DataTable();

            //columns              
            table.Columns.Add("ClientId", typeof(long));
            table.Columns.Add("Cpf", typeof(string));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Phone", typeof(string));
            table.Columns.Add("Email", typeof(string));

            //rows  
            foreach (var item in clientes)
            {
                table.Rows.Add(
                    item.ClientId,
                    item.Cpf,
                    item.Name,
                    item.Phone,
                    item.Email);
            }

            return table;
        }

        public static void SaveCsvFile(this DataTable dataTable, string strFilePath)
        {
            var lines = new List<string>();

            string[] columnNames = dataTable.Columns
                .Cast<DataColumn>()
                .Select(column => column.ColumnName)
                .ToArray();

            var header = string.Join(",", columnNames.Select(name => $"\"{name}\""));
            lines.Add(header);

            var valueLines = dataTable.AsEnumerable()
                .Select(row => string.Join(",", row.ItemArray.Select(val => $"\"{val}\"")));

            lines.AddRange(valueLines);

            File.WriteAllLines(strFilePath, lines, Encoding.UTF8);
        }

        public static void SaveExcelByInterop(this DataTable dataTable, FileInfo file)
        {
            if (dataTable == null || dataTable.Columns.Count == 0)
                return;

            DeleteFile(file);

            var excelApp = new Excel.Application();
            excelApp.Workbooks.Add();

            Excel._Worksheet workSheet = excelApp.ActiveSheet;

            // column headings
            for (var i = 0; i < dataTable.Columns.Count; i++)
            {
                workSheet.Cells[1, i + 1] = dataTable.Columns[i].ColumnName;
            }

            // rows
            for (var i = 0; i < dataTable.Rows.Count; i++)
            {
                // to do: format datetime values before printing
                for (var j = 0; j < dataTable.Columns.Count; j++)
                {
                    workSheet.Cells[i + 2, j + 1] = dataTable.Rows[i][j];
                }
            }

            excelApp.Columns.AutoFit();

            if (!string.IsNullOrEmpty(file.FullName))
            {
                workSheet.SaveAs(file.FullName);
                excelApp.Quit();
            }
            else
            {
                excelApp.Visible = true;
            }
        }

        public static void SaveExcelByClosedXml(this DataTable dataTable, string path)
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.Add(dataTable, "Client");
                ws.Columns().AdjustToContents();
                ws.Table(0).Theme = XLTableTheme.None;                            
                wb.SaveAs(@$"{path}\Client.xlsx");
            }
        }

        public static void DeleteFile(FileInfo file)
        {
            if (file.Exists)
                file.Delete();
        }
    }
}
