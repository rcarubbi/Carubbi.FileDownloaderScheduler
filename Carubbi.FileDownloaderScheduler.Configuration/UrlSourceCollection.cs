using System.Configuration;

namespace Carubbi.FileDownloaderScheduler.Configuration
{

    public class UrlSourceCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new UrlSourceElement();
        }
        
        
        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((UrlSourceElement)element).Path;
        }
    }
}


