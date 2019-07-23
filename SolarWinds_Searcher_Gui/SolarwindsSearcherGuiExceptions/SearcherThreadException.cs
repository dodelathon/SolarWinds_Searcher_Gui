using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomExceptions
{
    class SearcherThreadException : Exception
    {
        public SearcherThreadException()
        {

        }

        public SearcherThreadException(string message) : base(string.Format("{0} An error within Thread {0} has caused the thread to cease!", message))
        {

        }

        public SearcherThreadException(string message, Exception inner) : base(message, inner)
        {

        }
    }
}
