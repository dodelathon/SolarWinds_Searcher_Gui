using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomExceptions
{
    public class CantOpenException : ExcelInterationException
    {
        public CantOpenException()
        {

        }

        public CantOpenException(string Message) : base (Message)
        {

        }

        public CantOpenException(string Message, Exception inner) : base(Message, inner)
        {

        }
    }
}
