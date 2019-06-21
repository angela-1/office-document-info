using McMaster.Extensions.CommandLineUtils;
using System;
using System.Collections.Generic;
using System.IO;

namespace office_document_info
{
    class Program
    {
        static int Main(string[] args)
        {
            return CommandLineApplication.Execute<Program>(args);
        }

        [Argument(0, Description = "Full path of docx file or folder")]
        public string FilePath { get; }

        [Option(ShortName = "f", LongName = "format", Description = "Output format, json or line")]
        public string Format { get; }


        private List<string> FilterDocx(List<string> filepaths)
        {
            List<string> result = new List<string>();
            foreach (var item in filepaths)
            {
                String ext = Path.GetExtension(item);
                if (ext == ".docx")
                {
                    result.Add(item);
                }
            }
            return result;
        }

        private void OnExecute()
        {
            bool v = FilePath == null;
            if (v)
            {
                Console.WriteLine("No input, please input file or folder");
                return;
            }
            bool isFolder = Directory.Exists(FilePath);
            string formatString = Format ?? "line";
            OutputFormat outputFormat = formatString == "json" ? OutputFormat.JSON : OutputFormat.LINE;
            if (isFolder)
            {
                string[] filePaths = Directory.GetFiles(FilePath);
                List<string> docxFiles = FilterDocx(new List<string>(filePaths));
                foreach (var docx in docxFiles)
                {
                    var doc = new DocumentInfo();
                    var result = doc.GetInfo(docx, outputFormat);
                    Console.WriteLine(result);
                }
            }
            else
            {
                string ext = Path.GetExtension(FilePath);
                if (ext != ".docx")
                {
                    Console.WriteLine($"Not support {ext} format");
                    return;
                }
                var doc = new DocumentInfo();
                var result = doc.GetInfo(FilePath, outputFormat);
                Console.WriteLine(result);
            }
        }
    }
}
