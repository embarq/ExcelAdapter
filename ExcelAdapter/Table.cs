using System.Collections.Generic;

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

        public static Table Import(List<List<KeyValuePair<string, string>>> importData)
        {
            List<string> fields = new List<string>();
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
    }
}
