using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;

namespace Excalibur
{
    public interface IWebRequestFactory
    {
        HttpWebRequest Create(string url);
    }

    public class HttpWebRequestFactory : IWebRequestFactory
    {
        public HttpWebRequest Create(string url)
        {
            return (HttpWebRequest)WebRequest.Create(url);
        }
    }

}
