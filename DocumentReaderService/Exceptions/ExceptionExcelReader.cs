using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentReaderService.Exceptions
{
    public class ExceptionExcelReader : Exception
    {
        public ExceptionExcelReader()
        {
        }

        public ExceptionExcelReader(string message)
            :base(message)
        {
        }
    }
}
