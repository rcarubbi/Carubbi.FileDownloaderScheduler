using System.Configuration;

namespace Carubbi.FileDownloaderScheduler.Configuration
{
    public class UrlSources : ConfigurationSection
    {
        [ConfigurationProperty("", IsRequired = true, IsDefaultCollection = true)]
        public UrlSourceCollection Instances
        {
            get { return (UrlSourceCollection)this[""]; }
            set { this[""] = value; }
        }
    }
}
