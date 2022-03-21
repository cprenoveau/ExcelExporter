# ExcelExporter
Simple C# Excel to JSON exporter

Usage from within Unity Editor:

        public static void Convert(bool hidden)
        {
            var executablePath = Path.Combine(Application.dataPath, @"..\ExcelExporter.exe");
            var inputPath = Path.Combine(Application.dataPath, "Excel");
            var outputPath = Path.Combine(Application.dataPath, "Resources");

            ProcessStartInfo processInfo = new ProcessStartInfo();

            processInfo.UseShellExecute = true;
            if (hidden) processInfo.WindowStyle = ProcessWindowStyle.Hidden;
            processInfo.FileName = executablePath;
            processInfo.Arguments = inputPath +" "+ outputPath;

            var process = Process.Start(processInfo);
            process.WaitForExit();
        }
