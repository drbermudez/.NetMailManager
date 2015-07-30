using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NetMail;

namespace MailApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                MailManager mail = new MailManager();

                mail.Body = txtMessage.Text;
                mail.Subject = txtSubject.Text;
                mail.Receiver = txtTo.Text;
                mail.ReceiverDisplayName = "David R Bermudez";
                mail.SMTPPort = 587;
                mail.SMTPUsername = "eligibilityerror@firstmedicalpr.com";
                mail.SMTPHost = "smtp.office365.com";
                mail.SMTPPassword = "R3p0rtEg1";
                mail.EnableSSL = true;
                mail.Priority = System.Net.Mail.MailPriority.High;
                mail.Sender = "eligibilityerrorreporting@firstmedicalpr.com";
                mail.SenderDisplayName = "Mr Microsoft";
                mail.AttachmentUrls = new List<string>();
                bool seFue = mail.SendMail();

                if (!seFue)
                {
                    foreach (ErrorManager error in mail.Errors)
                    {
                        if (error.InnerException != null)
                        {
                            MessageBox.Show(error.Message + ": " + error.InnerException.Message);
                        }
                        else
                        {
                            MessageBox.Show(error.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Not implemented");
        }
    }
}
