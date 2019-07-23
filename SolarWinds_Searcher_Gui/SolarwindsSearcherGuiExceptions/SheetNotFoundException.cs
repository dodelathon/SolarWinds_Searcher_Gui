using System;


namespace CustomExceptions
{
    public class SheetNotFoundException : ExcelInterationException
    {
        public SheetNotFoundException()
        {

        }
       
        public SheetNotFoundException(string message) : base(string.Format("{0} was unable to be found within the workbook!", message))
        {

        }

        public SheetNotFoundException(string message, Exception inner) : base(message, inner)
        {

        }
    }
}
