using System.Collections.Specialized;

namespace Carubbi.FileDownloaderScheduler.Configuration
{
    public class UrlSourcesConfig
    {
        protected static StringCollection _paths ;
        
        static UrlSourcesConfig()
        {
            _paths = new StringCollection();
            UrlSources sec = (UrlSources)System.Configuration.ConfigurationManager.GetSection("UrlSources");
            foreach (UrlSourceElement i in sec.Instances)
            {
                _paths.Add(i.Path);
            }
        }
        public static StringCollection Paths
        {
            get
            {
                return _paths;
            }
        }

        private UrlSourcesConfig()
        {
        }

    }



}
