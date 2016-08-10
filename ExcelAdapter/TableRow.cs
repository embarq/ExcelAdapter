using System.Collections.Generic;

namespace ExcelAdapter
{
    public class TableRow
    {
        private List<TableColumn> columns;

        public TableRow AddColumn(TableColumn tableColumn)
        {
            columns.Add(tableColumn);
            return this;
        }

        public TableRow AddColumn(string field, dynamic data)
        {
            columns.Add(new TableColumn(this.columns.Count, field, data));
            return this;
        }

        public List<TableColumn> GetColumns()
        {
            return columns;
        }

        /// <summary>
        /// Returns full row as string array
        /// </summary>
        /// <returns>string[]</returns>
        public string[] GetColumnsContent()
        {
            return new List<string>(System.Linq.Enumerable.Select<TableColumn, string>(columns, item => item.data)).ToArray();
        }

        public TableRow()
        {
            columns = new List<TableColumn>();
        }
    }
}
