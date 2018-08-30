using Carubbi.FileDownloaderScheduler.PluginInterfaces;
using Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace Carubbi.FileDownloaderScheduler.PluginXlsHtm
{
    public class PluginXslHtm : IFileDownloaderSchedulerPlugin
    {
        #region IFileDownloaderSchedulerPlugin Members

        private readonly string[] _arrCells =
        {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N",
            "O", "P", "Q", "R", "S", "T", "U", "V", "X", "W", "Y", "Z"
        };

        private readonly object _objMissing = Missing.Value;
        private Workbook _objWorkBook;
        private Worksheet _objWorkSheet;
        private ApplicationClass _objExcelApp;


        public string GetHtmlPath(string caminho)
        {
            var caminhoTemporario = Path.Combine(Path.GetDirectoryName(caminho) ?? throw new InvalidOperationException(),
                $"tmp{Path.GetFileNameWithoutExtension(caminho)}.{Path.GetExtension(caminho)}");
            File.Copy(caminho, caminhoTemporario);

            _objExcelApp = new ApplicationClass();
            _objWorkBook = _objExcelApp.Workbooks.Open(caminho, _objMissing, _objMissing, _objMissing, _objMissing,
                _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing,
                _objMissing, _objMissing);

            _objWorkSheet = (Worksheet) _objWorkBook.Worksheets[1];

            var rangeData = (Range) _objWorkSheet.Cells[2, _arrCells[13]];
            var rangeTipoRelatorio = (Range) _objWorkSheet.Cells[2, _arrCells[16]];

            var data = rangeData.Value2.ToString();
            var mes = ConvertToMonthNumber(data.Substring(0, 3));
            var ano = data.Substring(4, 2);
            var tipoXls = rangeTipoRelatorio.Value2.ToString();

            _objWorkBook.Close(false, caminhoTemporario, _objMissing);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(rangeData);
            Marshal.ReleaseComObject(rangeTipoRelatorio);
            Marshal.ReleaseComObject(_objWorkSheet);
            Marshal.ReleaseComObject(_objWorkBook);
            _objExcelApp.Quit();
            Marshal.ReleaseComObject(_objExcelApp);


            File.Delete(caminhoTemporario);

            var htmlPath = Path.Combine(Path.GetDirectoryName(caminho) ?? throw new InvalidOperationException(),
                $"{Path.GetFileNameWithoutExtension(caminho)}_{mes}_{ano} ({tipoXls}).htm");

            return htmlPath;
        }

        private static string ConvertToMonthNumber(string monthName)
        {
            switch (monthName.ToLower())
            {
                case "jan":
                    return "01";
                case "fev":
                    return "02";
                case "mar":
                    return "03";
                case "abr":
                    return "04";
                case "mai":
                    return "05";
                case "jun":
                    return "06";
                case "jul":
                    return "07";
                case "ago":
                    return "08";
                case "set":
                    return "09";
                case "out":
                    return "10";
                case "nov":
                    return "11";
                case "dez":
                    return "12";
                default:
                    return "";
            }
        }

        public PluginXslHtm()
        {
            Name = "Plugin de conversão de xls em html";
        }

        private readonly string _excelPath = ConfigurationManager.AppSettings["excelPath"];
 

        public List<KeyValuePair<string, Stream>> Process(KeyValuePair<string, Stream> input)
        {
            if (!input.Key.EndsWith(".xlsm")) return null;
            var tempDirectory = Path.Combine(Environment.CurrentDirectory, "temp");

            Directory.Delete(tempDirectory, true);
            Directory.CreateDirectory(tempDirectory);

            var fullpath = Path.Combine(tempDirectory, Path.GetFileName(input.Key));

            var retorno = new List<KeyValuePair<string, Stream>>();
            var fs = new FileStream(fullpath, FileMode.Create);

            input.Value.CopyTo(fs);
            fs.Close();

            var proc = new Process
            {
                StartInfo =
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    FileName = Path.Combine(_excelPath, "excel.exe"),
                    Arguments = "\"" + fullpath + "\""
                }
            };
            proc.Start();

            Thread.Sleep(Convert.ToInt32(ConfigurationManager.AppSettings["secondsWaitProcess"]) * 1000);
            if (!proc.HasExited)
                proc.Kill();

            if (proc.StandardError.ReadToEnd().Length > 0)
                return null;


            var htmlFullPath = GetHtmlPath(fullpath);

            var fsHtml = File.OpenRead(htmlFullPath);

            retorno.Add(new KeyValuePair<string, Stream>(htmlFullPath, fsHtml));
            var folderFullPath = Path.Combine(Path.GetDirectoryName(htmlFullPath) ?? throw new InvalidOperationException(),
                string.Concat(Path.GetFileNameWithoutExtension(htmlFullPath), "_arquivos"));

            var files = Directory.GetFiles(folderFullPath);
            retorno.AddRange(from file in files let fsFile = File.OpenRead(file) select new KeyValuePair<string, Stream>(file, fsFile));


            return retorno;

        }

        public string Name { get; }

        #endregion
    }
}