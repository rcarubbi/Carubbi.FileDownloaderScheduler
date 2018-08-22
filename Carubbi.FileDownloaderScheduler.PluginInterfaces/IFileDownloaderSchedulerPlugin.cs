using System.Collections.Generic;
using System.IO;

namespace Carubbi.FileDownloaderScheduler.PluginInterfaces
{
    public interface IFileDownloaderSchedulerPlugin
    {
        string Name { get; }
        List<KeyValuePair<string, Stream>> Process(KeyValuePair<string, Stream> input);
    }
}