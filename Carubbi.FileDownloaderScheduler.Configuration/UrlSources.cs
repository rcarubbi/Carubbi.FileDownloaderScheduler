using System.Configuration;

namespace Carubbi.FileDownloaderScheduler.Configuration
{
    public class UrlSources : ConfigurationSection
    {
        [ConfigurationProperty("", IsRequired = true, IsDefaultCollection = true)]
        public UrlSourceCollection Instances
        {
            get => (UrlSourceCollection) this[""];
            set => this[""] = value;
        }
    }
}