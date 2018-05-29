﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace DocxEmailToWordPress
{
    class SendEmail
    {

        public string username = "***REMOVED***";
        public string Password = "***REMOVED***";
        public String _smtpHost;
        public String _sendTo;
        public String _sentFrom;
        public int _port;
        public String bcc = String.Empty;
        public String cc = String.Empty;

        public SendEmail(String smtpHost, String sendTo, String sentFrom, int port)
        {
            _smtpHost = smtpHost;
            _sendTo = sendTo;
            _sentFrom = sentFrom;
            _port = port;

        }


        public Boolean Send(List<PostLog> postLog)
            {
            
            SmtpClient smtpClient = new SmtpClient();

            NetworkCredential basicCredential = new NetworkCredential(username, Password);
                MailMessage message = new MailMessage();
                message.From = new MailAddress(_sentFrom);
                message.To.Add(new MailAddress(_sendTo));
                

                smtpClient.Host = _smtpHost;
                smtpClient.Port = _port;
               // smtpClient.UseDefaultCredentials = false;
                smtpClient.EnableSsl = true;
                smtpClient.Credentials = basicCredential;
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
               

            

            // Check if the bcc value is null or an empty string
            if ((bcc != null) && (bcc != string.Empty))
                {
                    // Set the Bcc address of the mail message
                    message.Bcc.Add(new MailAddress(bcc));
                }      // Check if the cc value is null or an empty value
                if ((cc != null) && (cc != string.Empty))
                {
                    // Set the CC address of the mail message
                    message.CC.Add(new MailAddress(cc));
                }       // Set the subject of the mail message

                message.Subject = "Schools Posted";

                //Set IsBodyHtml to true means you can send HTML email.
                message.IsBodyHtml = true;

                // Set the priority of the mail message to normal
                message.Priority = MailPriority.Normal;

            message.Body = BuildMessage(postLog);

                // Set the method that is called back when the send operation ends.
                smtpClient.SendCompleted += new SendCompletedEventHandler(SendCompletedCallback);


            try
                {


                object userState = message;


                smtpClient.Send(message);


                Console.WriteLine("Sending message..");

                

                return true;

                }
                catch (Exception ex)
                {
                    //Error, could not send the message
                    Console.Write(ex.Message);


                return false;
                }

            

        }

        public String BuildMessage(List<PostLog> p)
        {
            StringBuilder SbAttachments = new StringBuilder();
            StringBuilder SbErrorMessages = new StringBuilder();
            String htmlString = String.Empty;

            foreach (var log in p)
            {

                var postData = log.PostStatus;
                var emailFrom = log.FromAddress;
                var sentTo = log.ToAddress;
                var timeRecieved = log.TimeRecieved;
                var subject = log.Subject;
                var logAttachments = SbAttachments.ToString();
                var posted = log.Posted.ToString();
                var logMessages = log.Messages;
                var errorMessages = SbErrorMessages.ToString();
                var htmlTable = log.PostedHtml;

                foreach (var attachment in log.Attachments)
                {
                    SbAttachments.Append("Attechment is " + attachment.Key.ToString() + "File size is: " + attachment.Value.ToString());


                }

                foreach (var logMessage in logMessages)
                {
                    SbErrorMessages.Append(logMessage.ToString());
                }

                HtmlString html = new HtmlString($"<body><p>&nbsp;</ p><table width =\"680\" border=\"1\" cellpadding=\"1\"><tr><td width=\"172\">Post Status:</td><td width=\"492\">{postData}</td></tr><tr><td>Email From:</td><td>{emailFrom}</td></tr><tr><td height=\"33\">Sent To:</td><td>{sentTo}</td></tr><tr><td>Time Recieved:</td><td>{timeRecieved}</td></tr><tr><td>Subject:</td><td>{subject}</td></tr><tr><td>Attactments:</td><td>{logAttachments}</td></tr><tr><td>Posted?:</td><td>{posted}</td></tr><tr><td>Error Messages:</td><td>{errorMessages}</td></tr></table><p>{htmlTable}</p><p>&nbsp;</p><p>&nbsp;</p><p>_______________________________________________________________________________________________</p><p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p></body>");

                htmlString = html.ToString();
            }


            return htmlString;

        } // end sendMail


        
        // async call back #not in use#

        public void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
           

            //Get the Original MailMessage object
            MailMessage mail = (MailMessage)e.UserState;

            //write out the subject
            string subject = mail.Subject;

            if (e.Cancelled)
            {
                Console.WriteLine("[{0}] Send canceled.", subject);
            }
            if (e.Error != null)
            {
                Console.WriteLine("[{0}] {1}", subject, e.Error.ToString());
            }
            else
            {
                Console.WriteLine("Message [{0}] sent.", subject);
               
            }
            

        }


    }
    
}


