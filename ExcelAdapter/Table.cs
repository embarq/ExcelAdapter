using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAdapter
{
    public class Table
    {
        public string[] fields;
        public List<TableRow> rows;

        /// <summary>Add row to the current table</summary>
        /// <returns>Ref to the last row in rows</returns
        public TableRow AddRow(List<KeyValuePair<string, string>> row)
        {
            TableRow newRow = new TableRow();
            foreach (KeyValuePair<string, string> col in row)
            {
                newRow.AddColumn(col.Key, col.Value);
            }
            rows.Add(newRow);
            return newRow;
        }

        public TableRow GetRow(int id)
        {
            return rows[id];
        }

        /// <param name="name">Table name</param>
        /// <param name="fields">Heading labels</param>
        public Table(string[] fields)
        {
            this.fields = fields;
            this.rows = new List<TableRow>();
        }

        public static Table Import(string path)
        {
            var importData = JsonImport.From(path);
            var fields = new List<string>();
            foreach (var data in importData[0])
            {
                fields.Add(data.Key);
            }

            Table table = new Table(fields.ToArray());

            foreach (var data in importData)
            {
                table.AddRow(data);
            }

            return table;
        }

        public void ExportExcel()
        {
            string filename = @"/database.xlsx";
            string currentPath = Directory.GetCurrentDirectory();
            string excelPath = currentPath + filename;

            char[] alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
            char left = alpha[0];
            char right = alpha[fields.Length - 1];

            var missing = System.Type.Missing;
            var excel = new Excel.Application();
            var workbook = excel.Workbooks.Add(missing);
            var workSheet = (Excel.Worksheet) (workbook.Worksheets[1]);

            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.Workbooks.Add(missing);

            workSheet.Range[left + "1", right + "1"].Value = fields;

            int counter = 0;
            while (counter < rows.Count)
            {
                int padding = counter + 2;
                string a = left + padding.ToString();
                string b = right + padding.ToString();
                workSheet.Range[a, b].Value = rows[counter].GetColumnsContent();
                counter++;
            }

            if (File.Exists(excelPath))
            {
                File.Delete(excelPath);
            }

            workSheet.SaveAs(
                excelPath,                              // file name
                Excel.XlFileFormat.xlWorkbookDefault,   // file format
                missing,                           // password
                missing,                           // write-reservation password
                false,                                  // read-only recommended
                false,                                  // create backup
                Excel.XlSaveAsAccessMode.xlNoChange,    // acces mode
                missing,                           // conflict resolution
                missing,                           // add workbook to the list of recent files
                missing);

            workbook.Close(missing, missing, missing);
            excel.Quit();
            excel = null;
        }
    }
}
