using System;
using System.Collections.Generic;
using System.Text;

namespace SheetReader
{
    public class TableBuilder : IEnumerable<ITableRow>
    {
        public string Separator { get; set; }

        protected List<ITableRow> rows = new List<ITableRow>();
        protected List<int> columnLength = new List<int>();

        public TableBuilder() : this("  ")
        {
            // Default separator " "
        }

        public TableBuilder(string separator)
        {
            Separator = separator;
        }

        public ITableRow AddRow(params object[] columns)
        {
            TableRow row = new TableRow(this);
            foreach (object column in columns)
            {
                string colStr = column.ToString().Trim();
                row.Add(colStr);

                if (row.Count > columnLength.Count)
                {
                    columnLength.Add(colStr.Length);
                }
                else
                {
                    if (colStr.Length > columnLength[row.Count - 1])
                    {
                        columnLength[row.Count - 1] = colStr.Length;
                    }
                }
            }

            rows.Add(row);
            return row;
        }

        protected string _formatString = null;
        public string FormatString
        {
            get
            {
                // generate a format string that 
                // accounts for the veriable lengths
                // of the columns and cache it.
                if (_formatString == null)
                {
                    string format = "";
                    int i = 0;
                    foreach (int length in columnLength)
                    {
                        // metaformat 
                        format += string.Format("{{{0},-{1}}}{2}", i++, length, Separator);
                    }

                    format += Environment.NewLine;
                    _formatString = format;
                }

                return _formatString;
            }
        }

        public override string ToString()
        {
            var strBld = new StringBuilder();
            foreach (TableRow row in rows)
            {
                row.ToFormatedString(strBld);
            }

            return strBld.ToString();
        }

        #region IEnumerable Members

        public IEnumerator<ITableRow> GetEnumerator()
        {
            return rows.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return rows.GetEnumerator();
        }

        #endregion
    }

}
