using System.Configuration;

namespace Carubbi.FileDownloaderScheduler.Configuration
{
    public class UrlSourceElement : ConfigurationElement
    {
        [ConfigurationProperty("path", IsRequired = true)]
        public string Path
        {
            get => (string) base["path"];
            set => base["path"] = value;
        }
    }
}