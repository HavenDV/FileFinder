using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading.Tasks;

[assembly: NeutralResourcesLanguage("en-US")]
namespace FileFinder
{
    class Program
    {
        //static DataSet LoadTextFile(string path)
        //{
        //    return new DataSet();
        //}

        static DataSet LoadExcelFile(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                var extension = Path.GetExtension(path);
                var isXls = string.Compare(extension, ".xls", StringComparison.OrdinalIgnoreCase) == 0;
                var isXlsx = string.Compare(extension, ".xlsx", StringComparison.OrdinalIgnoreCase) == 0;
                if (!isXls && !isXlsx)
                {
                    throw new ArgumentException(Resource.ExtensionException, path);
                }

                var reader = isXls ?
                    ExcelReaderFactory.CreateBinaryReader(stream) :
                    ExcelReaderFactory.CreateOpenXmlReader(stream);
                reader.IsFirstRowAsColumnNames = true;
                return reader.AsDataSet();
            }
        }

        static DateTime? ParseData(object text)
        {
            try
            {
                return (DateTime)text;
            }
            catch (InvalidCastException) {}

            string[] patterns = { "MM/dd/yyyy", "dd.MM.yy" };
            DateTime date;

            foreach (var pattern in patterns)
            {
                if (DateTime.TryParseExact((string)text, pattern, null,
                                          DateTimeStyles.None, out date))
                {
                    return date;
                }
            }
            //throw new FormatException(Resource.DateFormatException);
            return null;
        }

        static string FindFile(string startDir, string prefix)
        {
            Debug.WriteLine("FindFile. Current dir: {0}", startDir);
            foreach (var file in Directory.GetFiles(startDir)
                .Where(i => Path.GetFileName(i).StartsWith(prefix,StringComparison.OrdinalIgnoreCase)))
            {
                return file;
            }

            foreach (var dir in Directory.GetDirectories(startDir)
                .Where(i => string.Compare(
                    Path.GetFileName(i).Substring(0, 2),
                    prefix.Substring(0, 2),
                    StringComparison.OrdinalIgnoreCase) == 0)) {
                var file = FindFile(dir, prefix);
                if (file != null)
                {
                    return file;
                }
            }

            return null;
        }

        static void DisplayHelp()
        {
            Console.WriteLine(Resource.HelpText);
        }

        static bool HelpRequired(string param)
        {
            return param == "-h" || param == "--help" || param == "/?";
        }

        static void Main(string[] args)
        {
            try
            {
                if (args.Length < 1 || HelpRequired(args[0]))
                {
                    DisplayHelp();
                    return;
                }

                var filePath = args[0];
                var sourceDir = args.Length > 1 ? args[1] : ".";
                var destDir = args.Length > 2 ? args[2] : ".";
                var data = LoadExcelFile(filePath);
                Debug.Assert(data != null);
                Debug.Assert(data.Tables != null);
                Debug.Assert(data.Tables.Count > 0);
                var table = data.Tables[0];
                for (var i = 0; i < table.Rows.Count; ++i)
                {
                    var row = table.Rows[i];
                    var prefixStart = (string)row.ItemArray[0];
                    var date = ParseData(row.ItemArray[1]);
                    var dateString = date?.ToString("yyyyMMdd",null);
                    var prefix = (prefixStart ?? "") + "_" + (dateString ?? "");
                    var findedFile = FindFile(sourceDir, prefix);
                    if (findedFile != null)
                    {
                        var destFile = Path.Combine(destDir, Path.GetFileName(findedFile));
                        File.Move(findedFile, destFile);
                        Console.WriteLine("Result: {0}", findedFile);
                        Console.WriteLine("Prefix: {0}", prefix);
                        Console.WriteLine("Source Dir: {0}", sourceDir);
                        Console.WriteLine("Moved to: {0} as {1}", destDir, destFile);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(Resource.ErrorMessage, e.Message);
                //Console.WriteLine(Resource.StackTraceMessage);
                //Console.WriteLine(e.StackTrace);
            }
            //Console.ReadKey();
        }
    }
}
