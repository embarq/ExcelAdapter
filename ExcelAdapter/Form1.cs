using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;

namespace ExcelAdapter
{
    public partial class Form1 : Form
    {
        private Table table;
        private string defaultPath = Directory.GetCurrentDirectory() + @"/mock_data.json";

        public Form1()
        {
            InitializeComponent();
            InitializeContent(null);

            PrintSheet(table);
        }

        /// <summary>
        /// Initializing "table" instance
        /// </summary>
        /// <param name="path">Path to JSON-file</param>
        private void InitializeContent(string path)
        {
            bool isValidPath = !string.IsNullOrEmpty(path) && File.Exists(path);
            table = Table.Import(isValidPath ? path : defaultPath);
        }

        public void PrintSheet(Table table)
        {
            foreach (string heading in table.fields)
            {
                listView.Columns.Add(heading);
            }

            foreach (ColumnHeader columnHeader in listView.Columns)
            {
                columnHeader.Width = -2;
            }

            foreach (TableRow row in table.rows)
            {
                List<TableColumn> col = row.GetColumns();
                ListViewItem item = new ListViewItem(col[0].data);
                for (int i = 1; i < col.Count; i++)
                {
                    item.SubItems.Add(col[i].data);
                }
                listView.Items.Add(item);
            }
        }
    }
}
