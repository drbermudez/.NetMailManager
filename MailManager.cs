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
        /// Gets or Sets the sender's email address
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        public string Sender { get; set; }
        /// <summary>
        /// Gets or Sets the sender's display name
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
        /// Gets the messages that were downloaded through POP3 if fetching from the server
        /// </summary>
        public dsTables DownloadedMessages { get; private set; }

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
            DownloadedMessages = new dsTables();
            AttachmentUrls = new List<string>();
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
            smtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;

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

                client.TargetName = "STARTTLS/" + serverAddress;
                client.Send(mail);
                mail.Dispose();
                client.Dispose();

                return true;
            }
            catch (SmtpException smtp)
            {
                error = new ErrorManager();
                error.Message = smtp.Message;
                error.Source = smtp.Source;
                error.InnerException = smtp.InnerException;
                errorList.Add(error);

                return false;
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
                MailMessage mail = BuildMail(Sender, SenderDisplayName, Receiver, ReceiverDisplayName, Subject, Body, Priority, AttachmentUrls);

                client.TargetName = "STARTTLS/" + SMTPHost;
           
                client.Send(mail);
                mail.Dispose();
                client.Dispose();

                return true;
            }
            catch (SmtpException smtp)
            {
                error = new ErrorManager();
                error.Message = smtp.Message;
                error.Source = smtp.Source;
                error.InnerException = smtp.InnerException;
                errorList.Add(error);

                return false;
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
        /// Example showing:
        ///  - how to use UID's (unique ID's) of messages from the POP3 server
        ///  - how to download messages not seen before
        ///    (notice that the POP3 protocol cannot see if a message has been read on the server
        ///     before. Therefore the client need to maintain this state for itself)
        /// </summary>       
        /// <param name="seenUids">
        /// List of UID's of all messages seen before.
        /// New message UID's will be added to the list.
        /// Consider using a HashSet if you are using >= 3.5 .NET
        /// </param>
        /// <returns>A List of new Messages on the server</returns>
        public void FetchUnseenMessages(List<string> seenUids)
        {
            // The client disconnects from the server when being disposed
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                client.Connect(POPHost, POPPort, EnableSSL);

                // Authenticate ourselves towards the server
                client.Authenticate(POPUsername, POPPassword);

                // Fetch all the current uids seen
                List<string> uids = client.GetMessageUids();

                // Create a list we can return with all new messages
                List<Message> newMessages = new List<Message>();

                // All the new messages not seen by the POP3 client
                for (int i = 0; i < uids.Count; i++)
                {
                    string currentUidOnServer = uids[i];
                    if (!seenUids.Contains(currentUidOnServer))
                    {
                        // We have not seen this message before.
                        // Download it and add this new uid to seen uids

                        // the uids list is in messageNumber order - meaning that the first
                        // uid in the list has messageNumber of 1, and the second has 
                        // messageNumber 2. Therefore we can fetch the message using
                        // i + 1 since messageNumber should be in range [1, messageCount]
                        Message unseenMessage = client.GetMessage(i + 1);

                        // Add the message to the new messages
                        newMessages.Add(unseenMessage);

                        // Add the uid to the seen uids, as it has now been seen
                        seenUids.Add(currentUidOnServer);

                        PrepareMessageAndAttachments(newMessages);
                    }
                }                
            }            
        }

        /// <summary>
        /// Example showing:
        ///  - how to fetch all messages from a POP3 server
        /// </summary>
        /// <param name="hostname">Hostname of the server. For example: pop3.live.com</param>
        /// <param name="port">Host port to connect to. Normally: 110 for plain POP3, 995 for SSL POP3</param>
        /// <param name="useSsl">Whether or not to use SSL to connect to server</param>
        /// <param name="username">Username of the user on the server</param>
        /// <param name="password">Password of the user on the server</param>
        /// <returns>All Messages on the POP3 server</returns>
        public void FetchAllMessages()
        {
            // The client disconnects from the server when being disposed
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                client.Connect(POPHost, POPPort, EnableSSL);

                // Authenticate ourselves towards the server
                client.Authenticate(POPUsername, POPPassword);

                // Get the number of messages in the inbox
                int messageCount = client.GetMessageCount();

                // We want to download all messages
                List<Message> allMessages = new List<Message>(messageCount);

                // Messages are numbered in the interval: [1, messageCount]
                // Ergo: message numbers are 1-based.
                // Most servers give the latest message the highest number
                for (int i = messageCount; i > 0; i--)
                {
                    allMessages.Add(client.GetMessage(i));
                }

                PrepareMessageAndAttachments(allMessages);
            }
        }

        /// <summary>
        /// Prepares and inserts all messages and attachments into the corresponding data tables
        /// </summary>
        /// <param name="messageList">List of Message objects with messages</param>
        private void PrepareMessageAndAttachments(List<Message> messageList)
        {
            foreach (Message message in messageList)
            {
                dsTables.MessageRow row = DownloadedMessages.Message.NewMessageRow();

                row.Body = message.MessagePart.GetBodyAsText();
                row.DateSent = message.Headers.DateSent;
                row.FromEmail = message.Headers.From.Address;
                row.FromName = message.Headers.From.DisplayName;
                row.Importance = message.Headers.Importance.ToString();
                row.UniqueId = message.Headers.MessageId;
                row.Subject = message.Headers.Subject;
                for (int x = 0; x < message.Headers.To.Count; x++)
                {
                    row.ToAddr = message.Headers.To[x].Address + ";";
                    row.ToName = message.Headers.To[x].DisplayName + ";";
                }
                DownloadedMessages.Message.AddMessageRow(row);
                DownloadedMessages.Message.AcceptChanges();

                //try and get attachments
                int messageId = Int32.Parse(DownloadedMessages.Message.Rows[DownloadedMessages.Message.Rows.Count - 1]["MessageId"].ToString());
                List<MessagePart> attachments = message.FindAllAttachments();
                foreach (MessagePart attachment in attachments)
                {
                    dsTables.AttachmentsRow attach = DownloadedMessages.Attachments.NewAttachmentsRow();

                    if (attachment.IsAttachment)
                    {
                        attach.Attachment = attachment.Body;
                        attach.AttachmentName = attachment.FileName;
                        attach.MimeType = attachment.ContentType.MediaType;
                        attach.MessageId = messageId;
                        DownloadedMessages.Attachments.AddAttachmentsRow(attach);
                    }
                }
            }
        }
    }
}
