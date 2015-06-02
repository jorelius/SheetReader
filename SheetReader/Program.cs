using System;
using PowerArgs;

namespace SheetReader
{
    [TabCompletion, ArgDescription("SheetReader provides facilities for displaying excel workbooks and individual spreadsheets.")]
    public class SheetReaderOptions
    {
        [ArgRequired, ArgDescription("Path to Excel Workbook"), ArgShortcut("p"), ArgExistingFile]
        public string Path { get; set; }

        [ArgDescription("Excel Workbook sheet"), ArgShortcut("s")]
        public string Sheet { get; set; }

        [ArgDescription("List Excel Workbook sheets"), ArgShortcut("ss")]
        public bool ShowSheets { get; set; }

        [HelpHook, ArgDescription("Help and how to use"), ArgShortcut("h"), ArgShortcut("?")]
        public bool Help { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var options = Args.Parse<SheetReaderOptions>(args);

                // exit if help is requested. 
                if (options == null) return;

                using(var view = new ExcelTableView(options.Path))
                {
                    if (options.ShowSheets)
                    {
                        view.DisplayTableNames(); 
                    }
                    else
                    {
                        view.DisplayTable(options.Sheet);
                    }
                }

                Console.ReadLine();
            } 
            catch (ArgException m)
            {
                Console.WriteLine(m.Message);
                Console.WriteLine(ArgUsage.GenerateUsageFromTemplate<SheetReaderOptions>());
            }
        }
    }
}
