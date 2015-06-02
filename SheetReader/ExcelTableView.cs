using Excel;
using System;
using System.IO;

namespace SheetReader
{
    class ExcelTableView : IDisposable
    {
        private IExcelDataReader excelReader; 

        public ExcelTableView(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException();
            }

            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

            var extension = Path.GetExtension(filePath).ToLowerInvariant();

            switch(extension)
            {
                case ".xls" :
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                    break;

                case ".xlsx" :
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);;
                    break;
                default :
                    throw new ArgumentException();
            }
        }

        public void DisplayTableNames()
        {
            Console.WriteLine("Tables");
            Console.WriteLine("------");

            var result = excelReader.AsDataSet();
            for (int i = 0; i < result.Tables.Count; i++)
            {
                Console.WriteLine(result.Tables[i].TableName);
            }
        }

        /// <summary>
        /// Display table or tables in workbook
        /// </summary>
        /// <param name="tableName">Table to display. Can be empty and will show all tables.</param>
        public void DisplayTable(string tableName)
        {
            var result = excelReader.AsDataSet();
            for (int i = 0; i < result.Tables.Count; i++)
            {
                var table = result.Tables[i];
                if (string.IsNullOrWhiteSpace(tableName) || 
                    tableName.Trim().Equals(table.TableName.Trim()) )
                {
                    var builder = new TableBuilder();

                    for(int r = 0; r < table.Rows.Count; r++)
                    {
                        builder.AddRow(table.Rows[r].ItemArray);
                    }

                    Console.Write(builder.ToString());
                    Console.WriteLine();
                }
            }
        }

        public void Dispose()
        {
            if (excelReader != null)
            {
                excelReader.Dispose(); 
            }
        }
    }
}
