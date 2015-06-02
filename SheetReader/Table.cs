namespace SheetReader
{
    public class Table
    {
        public string TableName { get; private set; }

        public int ColumnCount { get; private set; }

        public int RowCount { get; private set; }

        public string[,] TableData;

        public int[] MaxColumnSizes; 

        public Table(System.Data.DataTable table)
        {
            TableName = table.TableName;

            ColumnCount = table.Columns.Count;
            RowCount = table.Rows.Count;
            MaxColumnSizes = new int[ColumnCount];

            // fill table data and the maximum character 
            // count
            TableData = new string[RowCount, ColumnCount];
            for(int i = 0; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i]; 
                for(int j = 0; j < row.ItemArray.Length; j++)
                {
                    TableData[i,j] = row.ItemArray[j].ToString();

                    if (TableData[i, j].Length > MaxColumnSizes[j])
                    {
                        MaxColumnSizes[j] = TableData[i, j].Length;
                    }
                }
            }
        }
    }
}
