using System.Collections.Specialized;
using System.Configuration;

namespace Carubbi.FileDownloaderScheduler.Configuration
{
    public class UrlSourcesConfig
    {
        protected static StringCollection _paths;

        static UrlSourcesConfig()
        {
            _paths = new StringCollection();
            var sec = (UrlSources) ConfigurationManager.GetSection("UrlSources");
            foreach (UrlSourceElement i in sec.Instances) _paths.Add(i.Path);
        }

        private UrlSourcesConfig()
        {
        }

        public static StringCollection Paths => _paths;
    }
}