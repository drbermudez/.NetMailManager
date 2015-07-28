using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Data;
using System.Net.Mail;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Net.Security;
using OpenPop.Common;
using OpenPop.Mime;
using OpenPop.Pop3;
using NetMail.DataSets;

namespace NetMail
{
    public class MailManager
    {
        /// <summary>
        /// Gets or sets the message's delivery priority
        /// </summary>
        public MailPriority Priority { get; set; }
        /// <summary>
        /// Gets or Sets the Email's subject
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public string Subject { get; set; }
        /// <summary>
        /// Gets or Sets the Email's body
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public string Body { get; set; }
        /// <summary>
        /// Gets or Sets the recipient's email address
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public string Receiver { get; set; }
        /// <summary>
        /// Gets or Sets the recipient's display name
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public string ReceiverDisplayName { get; set; }
        /// <summary>
        /// Gets or Sets the sender's's display name
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public string SenderDisplayName { get; set; }
        /// <summary>
        /// Gets or Sets the Email's attachment url
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public List<string> AttachmentUrls { get; set; }
        /// <summary>
        /// Gets or Sets the smtp server's username, usually the user's email address
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public string SMTPUsername { get; set; }
        /// <summary>
        /// Gets or Sets the smtp servers password
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public string SMTPPassword { get; set; }
        /// <summary>
        /// Gets or Sets the smtp server's port
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public Int32 SMTPPort { get; set; }
        /// <summary>
        /// Gets or Sets the smtp server's address
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public string SMTPHost { get; set; }
        /// <summary>
        /// Gets or Sets the POP server's username
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public string POPUsername { get; set; }
        /// <summary>
        /// Gets or Sets the POP servers password
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public string POPPassword { get; set; }
        /// <summary>
        /// Gets or Sets the POP server's port
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public Int32 POPPort { get; set; }
        /// <summary>
        /// Gets or Sets the POP server's address
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public string POPHost { get; set; }
        /// <summary>
        /// Enable or disable SSL
        /// </summary>
        /// <value></value>
        /// <returns>True or False</returns>
        public bool EnableSSL { get; set; }
        /// <summary>
        /// Gets a list of errors that might have ocurred
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public List<ErrorManager> Errors
        {
            get { return errorList; }
        }

        /// <summary>
        /// Returns the mails from PRHIN
        /// </summary>
        /// <value>DataSet</value>
        /// <returns>DataSet</returns>
        public DataSet Messages
        {
            get { return theMail; }
        }

        private ErrorManager error;
        private List<ErrorManager> errorList;
        private DataSet theMail = new DataSet();

        /// <summary>
        /// Main constructor
        /// </summary>
        public MailManager()
        {
            error = null;
            errorList = new List<ErrorManager>();
            AttachmentUrls = null;
            EnableSSL = true;
        }

        /// <summary>
        /// Prepares the SMTP object to be used
        /// </summary>
        /// <param name="serverAddress">SMTP server IP or hostname</param>
        /// <param name="userName">Username</param>
        /// <param name="password">Password</param>
        /// <param name="port">SMTP port</param>
        /// <param name="useSSL">Use Secure Sockect Layer (true or false)</param>
        /// <returns>Pre-Built SMTP object</returns>
        private SmtpClient BuildSmtp(string serverAddress, string userName, string password, int port, bool useSSL)
        {
            SmtpClient smtpServer = new SmtpClient();

            smtpServer.UseDefaultCredentials = false;
            smtpServer.Credentials = new NetworkCredential(userName, password);
            smtpServer.Port = port;
            smtpServer.EnableSsl = useSSL;
            smtpServer.Host = serverAddress;

            return smtpServer;
        }

        /// <summary>
        /// Prepares the Mail object to be sent
        /// </summary>
        /// <param name="from">Sender email</param>
        /// <param name="fromDisplayName">Sender full name</param>
        /// <param name="to">Recipient email</param>
        /// <param name="toDisplayName">Recipient full name</param>
        /// <param name="subject">Subject</param>
        /// <param name="body">Message</param>
        /// <param name="priority">Message priority</param>
        /// <param name="attachmentsUrls">List of URLs or paths of the files to attach</param>
        /// <returns>Pre-Built Mail object</returns>
        private MailMessage BuildMail(string from, string fromDisplayName, string to, string toDisplayName, string subject, string body, MailPriority priority, List<string> attachmentsUrls)
        {
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress(from, fromDisplayName);
            mail.To.Add(new MailAddress(to, toDisplayName));
            if (attachmentsUrls.Count > 0)
            {
                foreach (string item in attachmentsUrls)
                {
                    mail.Attachments.Add(new Attachment(item));
                }
            }
            mail.Subject = Subject;
            mail.Body = Body;
            mail.Priority = priority;

            return mail;
        }

        /// <summary>
        /// Send a quick email message.  Use this method if not using the class properties
        /// </summary>
        /// <param name="from">Sender email address</param>
        /// <param name="fromDisplayName">Sender full name</param>
        /// <param name="to">Receiver email address</param>
        /// <param name="toDisplayName">Receiver full name</param>
        /// <param name="subject">Message subject</param>
        /// <param name="body">Message body</param>
        /// <param name="priority">Message priority</param>
        /// <param name="attachmentsUrls">Lisft of URLs or paths of the files to attach</param>
        /// <param name="useSSL">Use Secure Socket Layer (true or false)</param>
        /// <param name="serverAddress">Mail server's IP address or host name</param>
        /// <param name="userName">Sender's mail account username (usually the email address)</param>
        /// <param name="password">Sender's mail account password</param>
        /// <param name="port">SMTP port</param>
        /// <returns>True if successful; False if not</returns>
        public bool SendMail(string from, string fromDisplayName, string to, string toDisplayName, string subject,
                             string body, MailPriority priority, List<string> attachmentsUrls, bool useSSL,
                             string serverAddress, string userName, string password, int port)
        {
            try
            {
                SmtpClient client = BuildSmtp(serverAddress, userName, password, port, useSSL);
                MailMessage mail = BuildMail(from, fromDisplayName, to, toDisplayName, subject, body, priority, attachmentsUrls);

                client.Send(mail);
                mail.Dispose();
                client.Dispose();

                return true;
            }
            catch (Exception ex)
            {
                error = new ErrorManager();
                error.Message = ex.Message;
                error.Source = ex.Source;
                error.InnerException = ex.InnerException;
                errorList.Add(error);

                return false;
            }
        }

        /// <summary>
        /// Send a mail message.  Use this method with the class properties
        /// </summary>
        /// <returns>True if successful, otherwise False</returns>
        public bool SendMail()
        {
            try
            {
                SmtpClient client = BuildSmtp(SMTPHost, SMTPUsername, SMTPPassword, SMTPPort, EnableSSL);
                MailMessage mail = BuildMail(SMTPUsername, SenderDisplayName, Receiver, ReceiverDisplayName, Subject, Body, Priority, AttachmentUrls);

                client.Send(mail);
                mail.Dispose();
                client.Dispose();

                return true;
            }
            catch (Exception ex)
            {
                error = new ErrorManager();
                error.Message = ex.Message;
                error.Source = ex.Source;
                error.InnerException = ex.InnerException;
                errorList.Add(error);

                return false;
            }
        }
    }
}
