namespace ExcelAdapter
{
    public class TableColumn
    {
        public int row;
        public string field;
        public dynamic data;

        /// <param name="row">Link current `TableColumn` record to specific `TableRow`</param>
        /// <param name="field">Field name</param>
        public TableColumn(int row, string field, dynamic data)
        {
            this.row = row;
            this.field = field;
            this.data = data;
        }
    }
}
