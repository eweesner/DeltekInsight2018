using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net.Mail;
using Deltek.Framework.API.Server;
using Deltek.Vision.WorkflowAPI.Server;
using System.Data;

namespace ExpenseAlerts
{
    internal class VisionEmail : WorkflowBaseClass
    {
        private MailMessage mail;
        private SmtpClient client;
        struct CFGEmail
        {
            public string host;
            public string port;
            public string defaultSender;
            public string defaultRecipient;
            public string useSenderForReply;
            public string replyTo;
            public string prefixEmailSender;
            public string emailChunkSize;
            public string maxEmailSize;
        };

        public VisionEmail()
            : base()
        {
            string sql;
            sql = @"select * from CFGEmail";
            DataTable dt = new DataTable();
            try
            {
                dt = this.QueryData(sql);
            }
            catch (Exception e)
            {
                this.AddFatal(e.ToString() + "\r\n" + "VisionEmail contructor");
            }
            CFGEmail em;
            em.host = dt.Rows[0]["Host"].ToString();
            em.port = dt.Rows[0]["Port"].ToString();
            em.defaultSender = dt.Rows[0]["DefaultSender"].ToString();
            em.defaultRecipient = dt.Rows[0]["DefaultRecipient"].ToString(); //default help desk
            em.useSenderForReply = dt.Rows[0]["UseSenderForReply"].ToString();
            em.replyTo = dt.Rows[0]["ReplyTo"].ToString();
            em.prefixEmailSender = dt.Rows[0]["PrefixEmailSender"].ToString(); //add DeltekAdmin_ to sender email address
            em.emailChunkSize = dt.Rows[0]["EmailChunkSize"].ToString();
            em.maxEmailSize = dt.Rows[0]["MaxEmailSize"].ToString();


            mail = new MailMessage();
            client = new SmtpClient(em.host, Convert.ToInt32(em.port));
            //client.EnableSsl = true;
            if (em.prefixEmailSender.Equals("Y"))
                mail.From = new MailAddress("DeltekAdmin_" + em.defaultSender);
            if (em.useSenderForReply.Equals("N"))
                mail.ReplyToList.Add(em.defaultRecipient);
            
            mail.IsBodyHtml = true;
        }

        public void addTo(string address)
        {
            this.mail.To.Add(new MailAddress(address));
        }

        public void addCC(string address)
        {
            this.mail.CC.Add(new MailAddress(address));
        }

        public void addBcc(string address)
        {
            this.mail.Bcc.Add(new MailAddress(address));
        }

        public void subject(string subject)
        {
            this.mail.Subject = subject;
        }

        public void body(string body)
        {
            this.mail.Body = body;
        }

        public string sendMail()
        {
            string message = "";
            try
            {
                this.client.Send(this.mail);
                //this.AddInformation("Email sent successfully.");
            }
            catch (Exception e)
            {
                message = "Email did not send.  Back to the drawing board.\n\n" + e;
            }
            return message;
        }

        public string getAPPURL()
        { 
            string url = "";
            string sql;
            sql = "select AppURL from FW_CFGSystem";
            DataTable dt = new DataTable();
            try
            {
                dt = this.QueryData(sql);
            }
            catch (Exception e)
            {
                this.AddFatal(e.ToString() + "\r\n" + "VisionEmail.getAPPURL()");
            }
            url = dt.Rows[0]["AppURL"].ToString();
            return url;
        }

    }
}