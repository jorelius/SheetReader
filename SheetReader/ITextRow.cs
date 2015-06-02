using System;
using System.Text;

namespace SheetReader
{
    public interface ITableRow
    {
        void ToFormatedString(StringBuilder sb);
    }
}
