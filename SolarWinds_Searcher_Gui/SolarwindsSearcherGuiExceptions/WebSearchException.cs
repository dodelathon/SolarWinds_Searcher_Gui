using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomExceptions
{
    class WebSearchException : SearcherThreadException
    {
        public WebSearchException()
        {

        }

        public WebSearchException(string message) : base(string.Format("An error in searching on Thread: {0} has caused it to halt, results will be incomplete", message))
        {

        }

        public WebSearchException(Exception inner, string message) : base(message, inner)
        {

        }
    }
}
