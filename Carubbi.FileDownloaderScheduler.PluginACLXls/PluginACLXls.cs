using Carubbi.FileDownloaderScheduler.PluginInterfaces;
using Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Carubbi.FileDownloaderScheduler.PluginACLXls
{
    public class PluginACLXls : IFileDownloaderSchedulerPlugin
    {
        #region IFileDownloaderSchedulerPlugin Members

        public PluginACLXls()
        {
            Name = "Plugin do ACL";
        }

        private readonly string[] _arrCells =
        {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U",
            "V", "X", "W", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN",
            "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AX", "AW", "AY", "AZ"
        };

        private readonly object _objMissing = Missing.Value;
        private Workbook _objWorkBook;
        private Worksheet _objWorkSheet;
        private ApplicationClass _objExcelApp;
        private string _excelPath = ConfigurationManager.AppSettings["excelPath"];

     
        public List<KeyValuePair<string, Stream>> Process(KeyValuePair<string, Stream> input)
        {
            if (!input.Key.ToLower().EndsWith(".xls")) return null;
            var tempDirectory = Path.Combine(Environment.CurrentDirectory, "temp");

            if (Directory.Exists(tempDirectory)) Directory.Delete(tempDirectory, true);
            Directory.CreateDirectory(tempDirectory);

            var fullpath = Path.Combine(tempDirectory, Path.GetFileName(input.Key));
            var fs = new FileStream(fullpath, FileMode.Create);
            input.Value.CopyTo(fs);
            fs.Close();

            var dados = LerConteudo(fullpath);
            var xlsPath = Path.Combine(ConfigurationManager.AppSettings["xlsPath"], Path.GetFileName(input.Key));
            GravarConteudo(xlsPath, dados);

            FileStream fsOutput = null;
            fsOutput = File.OpenRead(xlsPath);


            var retorno = new List<KeyValuePair<string, Stream>>
            {
                new KeyValuePair<string, Stream>(Path.GetFileName(xlsPath), fsOutput)
            };

            return retorno;

        }

        private void GravarConteudo(string xlsPath, object[,] dados)
        {
            try
            {
                _objExcelApp = new ApplicationClass();
                _objWorkBook = _objExcelApp.Workbooks.Open(xlsPath, _objMissing, _objMissing, _objMissing, _objMissing,
                    _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing,
                    _objMissing,
                    _objMissing, _objMissing);


                _objWorkSheet = (Worksheet) _objWorkBook.Worksheets[2];

                for (var i = 1; i <= dados.GetLength(0); i++)
                for (var j = 1; j <= dados.GetLength(1); j++)
                    ((Range) _objWorkSheet.Cells[i, _arrCells[j - 1]]).Value2 = dados[i, j];

                _objWorkSheet = (Worksheet) _objWorkBook.Worksheets[1];
                ((PivotTable) _objWorkSheet.PivotTables(1)).RefreshTable();

                _objWorkBook.Save();
                _objWorkBook.Close(false, xlsPath, _objMissing);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(_objWorkSheet);
                Marshal.ReleaseComObject(_objWorkBook);
                _objExcelApp.Quit();
                Marshal.ReleaseComObject(_objExcelApp);
            }
        }

        private object[,] LerConteudo(string fullpath)
        {
            Range range = null;
            try
            {
                _objExcelApp = new ApplicationClass();
                _objWorkBook = _objExcelApp.Workbooks.Open(fullpath, _objMissing, _objMissing, _objMissing, _objMissing,
                    _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing,
                    _objMissing,
                    _objMissing, _objMissing);

                _objWorkSheet = (Worksheet) _objWorkBook.Worksheets[1];


                range = _objWorkSheet.Range["A1", Missing.Value];
                range = range.End[XlDirection.xlToRight];
                range = range.End[XlDirection.xlDown];
                var downAddress = range.Address[false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing];
                range = _objWorkSheet.Range["A1", downAddress];
                var values = (object[,]) range.Value2;

                _objWorkBook.Close(false, fullpath, _objMissing);

                return values;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(range);
                Marshal.ReleaseComObject(_objWorkSheet);
                Marshal.ReleaseComObject(_objWorkBook);
                _objExcelApp.Quit();
                Marshal.ReleaseComObject(_objExcelApp);
                File.Delete(fullpath);
            }
        }

        public string Name { get; }

        #endregion
    }
}