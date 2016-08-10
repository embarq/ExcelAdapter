using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace ExcelAdapter
{
    class Manager
    {
        public Table data { get; }
        private string defaultPath = @"./mock_data.json";
        public Manager(string path)
        {
            data = Table.Import(JSON(path ?? defaultPath));
            ExportExcel();
        }

        public List<List<KeyValuePair<string, string>>> JSON(string filename)
        {
            var data = new List<List<KeyValuePair<string, string>>>();
            dynamic jsonData = JArray.Parse(File.ReadAllText(filename));
            foreach (dynamic entry in jsonData)
            {
                var subSet = new List<KeyValuePair<string, string>>();
                foreach (dynamic property in entry)
                {
                    subSet.Add(new KeyValuePair<string, string>(property.Name, property.Value.ToString()));
                }
                data.Add(subSet);
            }
            return data;
        }

        public void ExportExcel()
        {
            string filename = @"/database.xlsx";
            string currentPath = Directory.GetCurrentDirectory();
            string excelPath = currentPath + filename;

            char[] alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
            char left = alpha[0];
            char right = alpha[data.fields.Length - 1];

            var excel = new Excel.Application();
            var workbook = excel.Workbooks.Add(Type.Missing);
            var workSheet = (Excel.Worksheet) (workbook.Worksheets[1]);

            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.Workbooks.Add(Type.Missing);

            workSheet.Range[left + "1", right + "1"].Value = data.fields;

            int counter = 0;
            while(counter < data.rows.Count) 
            {
                int padding = counter + 2;
                string a = left + padding.ToString();
                string b = right + padding.ToString();
                workSheet.Range[a, b].Value = data.rows[counter].GetColumnsContent();
                counter++;
            }

            if (File.Exists(excelPath))
            {
                File.Delete(excelPath);
            }

            workSheet.SaveAs(
                excelPath,                              // file name
                Excel.XlFileFormat.xlWorkbookDefault,   // file format
                Type.Missing,                           // password
                Type.Missing,                           // write-reservation password
                false,                                  // read-only recommended
                false,                                  // create backup
                Excel.XlSaveAsAccessMode.xlNoChange,    // acces mode
                Type.Missing,                           // conflict resolution
                Type.Missing,                           // add workbook to the list of recent files
                Type.Missing);

            workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            excel.Quit();
            excel = null;
        }
    }
}
