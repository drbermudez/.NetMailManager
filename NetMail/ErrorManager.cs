using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NetMail
{
    public class ErrorManager
    {
        public string Message { get; set; }
        public Exception InnerException { get; set; }
        public string Source { get; set; }

        public ErrorManager()
        {
            Message = "";
            InnerException = null;
            Source = "";
        }
    }
}
