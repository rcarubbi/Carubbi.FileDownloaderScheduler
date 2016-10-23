﻿using Carubbi.FileDownloaderScheduler.PluginInterfaces;
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

        private String[] _arrCells = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "X", "W", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AX", "AW", "AY", "AZ" };
        private object _objMissing = Missing.Value;
        private Workbook _objWorkBook;
        private Worksheet _objWorkSheet;
        private ApplicationClass _objExcelApp;
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
            if (input.Key.ToLower().EndsWith(".xls"))
            {
                string tempDirectory = Path.Combine(Environment.CurrentDirectory, "temp");

                if (Directory.Exists(tempDirectory))
                {
                    Directory.Delete(tempDirectory, true);
                }
                Directory.CreateDirectory(tempDirectory);

                string fullpath = Path.Combine(tempDirectory, Path.GetFileName(input.Key));
                FileStream fs = new FileStream(fullpath, FileMode.Create);
                CopyStream(input.Value, fs);
                fs.Close();

                object[,] dados = LerConteudo(fullpath);
                string xlsPath = Path.Combine(ConfigurationManager.AppSettings["xlsPath"], Path.GetFileName(input.Key));
                GravarConteudo(xlsPath, dados);

                FileStream fsOutput = null;
                fsOutput = File.OpenRead(xlsPath);


                List<KeyValuePair<string, Stream>> retorno = new List<KeyValuePair<string, Stream>>();
                retorno.Add(new KeyValuePair<String, Stream>(Path.GetFileName(xlsPath), fsOutput));

                return retorno;
            }
            return null;
        }

        private void GravarConteudo(string xlsPath, object[,] dados)
        {
            try
            {
                _objExcelApp = new ApplicationClass();
                _objWorkBook = _objExcelApp.Workbooks.Open(xlsPath, _objMissing, _objMissing, _objMissing, _objMissing,
                    _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing,
                    _objMissing, _objMissing);


                _objWorkSheet = (Worksheet)_objWorkBook.Worksheets[2];

                for (int i = 1; i <= dados.GetLength(0); i++)
                {
                    for (int j = 1; j <= dados.GetLength(1); j++)
                    {
                        ((Range)_objWorkSheet.Cells[i, _arrCells[j - 1]]).Value2 = dados[i, j];
                    }
                }

                _objWorkSheet = (Worksheet)_objWorkBook.Worksheets[1];
                ((PivotTable)_objWorkSheet.PivotTables(1)).RefreshTable();

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
                    _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing, _objMissing,
                    _objMissing, _objMissing);

                _objWorkSheet = (Worksheet)_objWorkBook.Worksheets[1];

               

                range = _objWorkSheet.get_Range("A1", Missing.Value);
                range = range.get_End(XlDirection.xlToRight);
                range = range.get_End(XlDirection.xlDown);
                string downAddress = range.get_Address(false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                range = _objWorkSheet.get_Range("A1", downAddress);
                object[,] values = (object[,])range.Value2;

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

        public string Name
        {
            get;
            private set;
        }

        #endregion
    }
}