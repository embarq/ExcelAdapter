using System.Collections.Generic;
using System.Windows.Forms;

namespace ExcelAdapter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            var manager = new Manager(null);
            PrintSheet(manager.data);
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
