using System.Collections.Generic;
using System.IO;

namespace Carubbi.FileDownloaderScheduler.PluginInterfaces
{
    public interface IFileDownloaderSchedulerPlugin
    {
        List<KeyValuePair<string, Stream>> Process(KeyValuePair<string, Stream> input);
        string Name { get; }
    }
}
