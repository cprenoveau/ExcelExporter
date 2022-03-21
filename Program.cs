using Newtonsoft.Json.Linq;
using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelExporter
{
    class Program
    {
        public static void Main(string[] args)
        {
            string inputPath = Directory.GetCurrentDirectory();
            string outputPath = Directory.GetCurrentDirectory();

            if (args.Length > 0)
                inputPath = args[0];

            if (args.Length > 1)
                outputPath = args[1];

            var inputDir = new DirectoryInfo(inputPath);

            Excel.Application excelApp = new Excel.Application();

            FileInfo[] files = inputDir.GetFiles("*.xlsx");
            foreach (var file in files)
            {
                if (file.Name[0] == '~')
                    continue;

                string excelFile = Path.GetFileNameWithoutExtension(file.Name);
                ExportExcelFile(excelApp, file.FullName, Path.Combine(outputPath, excelFile + ".json"));
            }

            excelApp.Quit();
        }

        private static void ExportExcelFile(Excel.Application excelApp, string excelFile, string outputPath)
        {
            Excel.Workbook book = null;

            try
            {
                Console.WriteLine("Exporting " + excelFile + " to " + outputPath);

                book = excelApp.Workbooks.Open(excelFile);
                Excel.Worksheet sheet = book.Worksheets.get_Item(1);

                File.WriteAllText(outputPath, ReadSheet(sheet, book).ToString());
            }
            catch (Exception e)
            {
                Console.Write("ExportExcelFile: Exception " + e.Message);
            }

            if (book != null)
                book.Close(false);
        }

        private static JToken ReadSheet(Excel.Worksheet sheet, Excel.Workbook book)
        {
            JArray jArray = new JArray();
            Excel.Range range = sheet.UsedRange;

            for (int r = 2; r <= range.Rows.Count; ++r)
            {
                JObject dict = new JObject();

                for (int c = 1; c <= range.Columns.Count; ++c)
                {
                    string label = (range.Cells[1, c] as Excel.Range).Value;
                    var value = (range.Cells[r, c] as Excel.Range).Value;

                    if (label[0] == '*')
                        continue;

                    var tokens = label.Split(':');
                    if (tokens.Length > 1 && tokens[0] == "sheet")
                    {
                        Excel.Worksheet subSheet = book.Worksheets[value];
                        dict.Add(tokens[1], ReadSheet(subSheet, book));
                    }
                    else
                    {
                        dict.Add(label, value);
                    }
                }

                jArray.Add(dict);
            }

            return jArray;
        }
    }
}
