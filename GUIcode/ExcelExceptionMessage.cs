using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComponentSearcher
{
    class ExcelExceptionMessage : Exception
    {
        
        public ExcelExceptionMessage()
        {
        }

        public ExcelExceptionMessage(string ExMsg) : base(ExMsg)
        {

        }

        public ExcelExceptionMessage(string ExMsg, Exception inner) : base(ExMsg, inner)
        {

        }

    }
}


