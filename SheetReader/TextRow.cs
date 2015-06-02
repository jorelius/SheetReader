using System;
using System.Collections.Generic;
using System.Text;

namespace SheetReader
{
    public class TableRow : List<string>, ITableRow
    {
        protected TableBuilder parent = null;
        public TableRow(TableBuilder Parent)
        {
            parent = Parent;
            if (parent == null)
            {
                throw new ArgumentException("Parent");
            }
        }

        public override string ToString()
        {
            var strBld = new StringBuilder();
            ToFormatedString(strBld);

            return strBld.ToString();
        }

        public void ToFormatedString(StringBuilder strBld)
        {
            strBld.AppendFormat(parent.FormatString, this.ToArray());
        }
    }
}
