using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomExceptions
{
    public class ExcelInterationException : Exception
    {
        public ExcelInterationException() : base()
        {

        }

        public ExcelInterationException(string message) : base (string.Format("{0} was the cause of this excepton", message))
        {

        }

        public ExcelInterationException(string message, Exception inner) : base (string.Format("{0} was the cause of this excepton", message), inner)
        {

        }

    }
}
