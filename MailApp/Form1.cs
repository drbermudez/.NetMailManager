﻿using System;
using System.Collections.Generic;
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
                mail.Recipients.Add(new Recipient(txtTo.Text, "Test Sender"));
                mail.SMTPPort = 465;
                mail.SMTPUsername = "info@transfieretuvinilpr.com"; 
                mail.SMTPHost = "smtpout.secureserver.net"; 
                mail.SMTPPassword = "Davi2002pr418219"; 
                mail.EnableSSL = true;
                mail.UseTLS = true;
                mail.Priority = System.Net.Mail.MailPriority.High;
                mail.Sender = "info@transfieretuvinilpr.com"; 
                mail.SenderDisplayName = "Mr Microsoft";
                mail.AttachmentUrls = new List<string>();
                bool seFue = mail.SendMail();

                if (!seFue)
                {
                    foreach (ErrorManager error in mail.Errors)
                    {
                        if (error.InnerException != null)
                        {
                            MessageBox.Show(error.Message + Environment.NewLine + error.InnerException.Message);
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

        private void Form1_Shown(object sender, EventArgs e)
        {
            txtTo.Focus();
        }
    }
}
