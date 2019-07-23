using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomExceptions
{
    class ChromeDriverStartUpException : SearcherThreadException
    {
        public ChromeDriverStartUpException()
        {

        }

        public ChromeDriverStartUpException(string message) : base(string.Format("Unable to start chromedriver: {0}", message))
        {

        }

        public ChromeDriverStartUpException(string message,Exception inner) : base( string.Format("Unable to start chromedriver: {0}", message), inner)
        {

        }


    }
}
