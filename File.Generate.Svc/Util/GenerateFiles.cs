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
            var columnsHeader = new Dictionary<string, Type>();

            columnsHeader.Add("ClientId", typeof(long));
            columnsHeader.Add("Cpf", typeof(string));
            columnsHeader.Add("Name", typeof(string));
            columnsHeader.Add("Phone", typeof(string));
            columnsHeader.Add("Email", typeof(string));

            var table = new DataTable();

            foreach (var item in columnsHeader)
            {
                table.Columns.Add(item.Key, item.Value);
            }

            foreach (var item in clientes)
            {
                table.Rows.Add(
                    item.ClientId,
                    item.Cpf,
                    Encoding.UTF8.GetString(Encoding.Default.GetBytes(item.Name)),
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

            for (var i = 0; i < dataTable.Columns.Count; i++)
            {
                workSheet.Cells[1, i + 1] = dataTable.Columns[i].ColumnName;
            }

            for (var i = 0; i < dataTable.Rows.Count; i++)
            {
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
            DeleteFile(new FileInfo(path));

            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add(dataTable, "Client");
            ws.Columns().AdjustToContents();
            ws.Table(0).Theme = XLTableTheme.None;
            wb.SaveAs(@$"{path}\Client.xlsx");
        }

        public static Stream SaveCsvStream(this DataTable dataTable)
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

            var csvBytes = Encoding.UTF8.GetBytes(string.Join(Environment.NewLine, lines));

            var csvMemoryStream = new MemoryStream(csvBytes);

            return csvMemoryStream;
        }

        public static Stream SaveExcelStream(this DataTable dataTable)
        {   
            using var wb = new XLWorkbook();
            
            IXLWorksheet ws = wb.Worksheets.Add(dataTable, "ClienteInfo");
            
            ws.Columns().AdjustToContents();

            ws.Table(0).Theme = XLTableTheme.None;                      
                
            var memoryStream = new MemoryStream();                      

            wb.SaveAs(memoryStream);

            //reset pointer position so the file is not inverted
            memoryStream.Seek(0, SeekOrigin.Begin);                            

            return memoryStream;
        }

        public static void DeleteFile(FileInfo file)
        {
            if (file.Exists)
                file.Delete();
        }
    }
}
