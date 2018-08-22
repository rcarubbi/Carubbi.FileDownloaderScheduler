using System.Collections.Generic;
using System.IO;
using Carubbi.FileDownloaderScheduler.PluginInterfaces;
using ICSharpCode.SharpZipLib.Zip;

namespace Carubbi.FileDownloaderScheduler.PluginZip
{
    public class PluginZip : IFileDownloaderSchedulerPlugin
    {
        #region IFileDownloaderSchedulerPlugin Members

        public PluginZip()
        {
            Name = "Plugin de Extração de arquivos Zip";
        }

        public string Name { get; }

        public List<KeyValuePair<string, Stream>> Process(KeyValuePair<string, Stream> input)
        {
            if (input.Key.EndsWith(".zip"))
            {
                var retorno = new List<KeyValuePair<string, Stream>>();
                using (var s = new ZipInputStream(input.Value))
                {
                    ZipEntry theEntry;

                    while ((theEntry = s.GetNextEntry()) != null)
                    {
                        var ms = new MemoryStream();
                        var size = 2048;
                        var data = new byte[2048];
                        while (true)
                        {
                            size = s.Read(data, 0, data.Length);
                            if (size > 0)
                                ms.Write(data, 0, size);
                            else
                                break;
                        }

                        ms.Position = 0;
                        retorno.Add(new KeyValuePair<string, Stream>(theEntry.Name, ms));
                    }
                }

                return retorno;
            }

            return null;
        }

        #endregion
    }
}