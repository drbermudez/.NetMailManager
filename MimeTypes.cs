using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace NetMail
{
    class MimeTypes
    {
        /// <summary>
        /// Returns the MIME type of the indicated file
        /// </summary>
        /// <param name="fileName">File name with extension</param>
        /// <returns>MIME type as string</returns>
        public static string GetMimeType(string fileName)
        {
            try
            {
                return System.Web.MimeMapping.GetMimeMapping(fileName);
            }
            catch
            {
                return "";
            }
        }
    }
}
