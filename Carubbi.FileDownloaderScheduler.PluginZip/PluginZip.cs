using Carubbi.FileDownloaderScheduler.PluginInterfaces;
using ICSharpCode.SharpZipLib.Zip;
using System.Collections.Generic;
using System.IO;

namespace Carubbi.FileDownloaderScheduler.PluginZip
{
    public class PluginZip : IFileDownloaderSchedulerPlugin
    {

        #region IFileDownloaderSchedulerPlugin Members

        public PluginZip()
        {
            Name = "Plugin de Extração de arquivos Zip"; 
        }

        public string Name
        {
            get;
            private set;
        }

        public List<KeyValuePair<string, Stream>> Process(KeyValuePair<string, Stream> input)
        {
           
            if (input.Key.EndsWith(".zip"))
            {
                List<KeyValuePair<string, Stream>> retorno = new List<KeyValuePair<string, Stream>>();
                using (ZipInputStream s = new ZipInputStream(input.Value))
                {

                    ZipEntry theEntry;

                    while ((theEntry = s.GetNextEntry()) != null)
                    {
                        MemoryStream ms = new MemoryStream();
                        int size = 2048;
                        byte[] data = new byte[2048];
                        while (true)
                        {
                            size = s.Read(data, 0, data.Length);
                            if (size > 0)
                            {
                                ms.Write(data, 0, size);
                            }
                            else
                            {
                                break;
                            }
                        }
                        ms.Position = 0;
                        retorno.Add(new KeyValuePair<string, Stream>(theEntry.Name, ms));
                    }
                }
                return retorno;
            }
            else
                return null;
        }
        #endregion
    }
}



