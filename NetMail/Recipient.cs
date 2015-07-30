using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NetMail
{
    public class Recipient
    {
        /// <summary>
        /// Email address
        /// </summary>
        public string Address { get; set; }
        /// <summary>
        /// Recipient's full display name
        /// </summary>
        public string FullName { get; set; }

        public Recipient(string email, string displayName)
        {
            Address = email;
            FullName = displayName;
        }
    }
}
