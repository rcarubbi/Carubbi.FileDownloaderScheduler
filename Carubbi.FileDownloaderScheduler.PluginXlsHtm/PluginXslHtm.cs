using Carubbi.FileDownloaderScheduler.PluginInterfaces;
using Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
namespace Carubbi.FileDownloaderScheduler.PluginXlsHtm
{
    public class PluginXslHtm : IFileDownloaderSchedulerPlugin
    {
        #region IFileDownloaderSchedulerPlugin Members


        private String[] _arrCells = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", 
            "O", "P", "Q", "R", "S", "T", "U", "V", "X", "W", "Y", "Z" };
        private object _objMissing = Missing.Value;
        private Workbook _objWorkBook;
        private Worksheet _objWorkSheet;
        private ApplicationClass _objExcelApp;


        public string GetHtmlPath(string caminho)
        {
            string htmlPath = string.Empty, tipoXls = string.Empty, mes = string.Empty, ano = string.Empty;

            string caminhoTemporario = Path.Combine(Path.GetDirectoryName(caminho), string.Format("tmp{0}.{1}", Path.GetFileNameWithoutExtension(caminho), Path.GetExtension(caminho)));
            File.Copy(caminho, caminhoTemporario);

            _objExcelApp = new ApplicationClass();
            _objWorkBook = _objExcelApp.Workbooks.Open(caminho, _objMissing, _objMissing, _objMissing, _objMissing,
                _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing,
                _objMissing, _objMissing);

            _objWorkSheet = (Worksheet)_objWorkBook.Worksheets[1];

            Range rangeData = (Range)_objWorkSheet.Cells[2, _arrCells[13]];
            Range rangeTipoRelatorio = (Range)_objWorkSheet.Cells[2, _arrCells[16]];

            string data = rangeData.Value2.ToString();
            mes = ConvertToMonthNumber(data.Substring(0, 3));
            ano = data.Substring(4, 2);
            tipoXls = rangeTipoRelatorio.Value2.ToString();

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
          
            htmlPath = Path.Combine(Path.GetDirectoryName(caminho), string.Format("{0}_{1}_{2} ({3}).{4}", Path.GetFileNameWithoutExtension(caminho), mes, ano, tipoXls, "htm"));

            return htmlPath;
        }
        
        private string ConvertToMonthNumber(string monthName)
        { 
            switch(monthName.ToLower())
            {
                case "jan" :
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

        private string _excelPath = ConfigurationManager.AppSettings["excelPath"].ToString();

        public static void CopyStream(Stream input, Stream output)
        {
            input.Position = 0;
            byte[] buffer = new byte[8 * 1024];
            int len;
            while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, len);
            }
        }

        public List<KeyValuePair<string, System.IO.Stream>> Process(KeyValuePair<string, System.IO.Stream> input)
        {
            if (input.Key.EndsWith(".xlsm"))
            {
                string tempDirectory = Path.Combine(Environment.CurrentDirectory, "temp");

                Directory.Delete(tempDirectory, true);
                Directory.CreateDirectory(tempDirectory);
                
                string fullpath = Path.Combine(tempDirectory,Path.GetFileName(input.Key));

                List<KeyValuePair<string, Stream>> retorno = new List<KeyValuePair<string, Stream>>();
                FileStream fs = new FileStream(fullpath, FileMode.Create);
                
                CopyStream(input.Value, fs);
                fs.Close();

                Process proc = new Process();
                proc.StartInfo.CreateNoWindow = true;
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.FileName = Path.Combine(_excelPath, "excel.exe");
                proc.StartInfo.Arguments = "\"" + fullpath + "\"";
                proc.Start();
                Thread.Sleep(Convert.ToInt32(ConfigurationManager.AppSettings["secondsWaitProcess"]) * 1000);
                if (!proc.HasExited)
                    proc.Kill();
               
                if (proc.StandardError.ReadToEnd().Length > 0)
                    return null;
                string folderFullPath = string.Empty;
                string htmlFullPath = string.Empty;
                FileStream fsHtml = null;


                htmlFullPath = GetHtmlPath(fullpath);

                fsHtml = File.OpenRead(htmlFullPath);

                retorno.Add(new KeyValuePair<String, Stream>(htmlFullPath, fsHtml));
                folderFullPath = Path.Combine(Path.GetDirectoryName(htmlFullPath), string.Concat(Path.GetFileNameWithoutExtension(htmlFullPath), "_arquivos"));
            
                string[] files = Directory.GetFiles(folderFullPath);
                foreach (string file in files)
                {
                    FileStream fsFile = File.OpenRead(file);
                    retorno.Add(new KeyValuePair<string,Stream>(file, fsFile));
                }


                return retorno;
            }
            return null;
        }

        public string Name
        {
            get;
            private set;
        }

        #endregion
    }
}
