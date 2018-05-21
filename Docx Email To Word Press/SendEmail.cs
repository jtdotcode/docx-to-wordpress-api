using System;
using System.Collections.Generic;
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
        public string _smtpHost { get; }
        public string _sendTo { get; }
        public string _sentFrom { get; }

        SendEmail(String smtpHost, String sendTo, String sentFrom)
        {
            smtpHost = _smtpHost;
            sendTo = _sendTo;
            sentFrom = _sentFrom;

        }
        
      

            public Boolean Send(List<PostLog> postLog)
            {

                
                SmtpClient smtpClient = new SmtpClient();
                NetworkCredential basicCredential = new NetworkCredential(username, Password);
                MailMessage message = new MailMessage();
                MailAddress fromAddress = new MailAddress(_sentFrom);

                smtpClient.Host = _smtpHost;
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = basicCredential;

                message.From = fromAddress;
                message.Subject = BuildMessage(postLog);
                //Set IsBodyHtml to true means you can send HTML email.
                message.IsBodyHtml = false;
                message.Body = body;
                message.To.Add(_sendTo);
                try
                {
                    
                smtpClient.Send(message);
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

            foreach (var log in p)
            {

                var postData = log.PostData;
                var emailFrom = log.FromAddress;
                var sentTo = log.ToAddress;
                var timeRecieved = log.TimeRecieved;
                var subject = log.Subject;
                var logAttachments = SbAttachments.ToString();
                var posted = log.Posted.ToString();
                var logMessages = log.Messages;
                var htmlTable = log.

                foreach (var attachment in log.Attachments)
                {
                    SbAttachments.Append("<td>" + "Attechment is " + attachment.Key + "File size is: " + attachment.Value.ToString() + "</td>");


                }

                foreach (var logMessage in logMessages)
                {
                    SbErrorMessages.Append("<td>" + logMessage + "</td>");
                }

                HtmlString htmlString = new HtmlString($"< body >< p > &nbsp;</ p >< table width =\"680\" border=\"1\" cellpadding=\"1\"><tr><td width=\"172\">Post Data</td><td width=\"492\">{postData}</td></tr><tr><td>Email From</td><td>{emailFrom}</td></tr><tr><td height=\"33\">Sent To</td><td>{sentTo}</td></tr><tr><td>Time Recieved</td><td>{timeRecieved}</td></tr><tr><td>Subject</td><td>{subject}</td></tr><tr><td>Attactments</td><td>{logAttachments}</td></tr><tr><td>Posted?</td><td>{posted}</td></tr><tr><td>Error Messages</td><td>{logMessages}</td></tr></table><p>{htmlTable}</p><p>&nbsp;</p><p>&nbsp;</p><p>_______________________________________________________________________________________________</p><p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p></body>");

            }



            

            




            


            return htmlString.ToString();

        } // end sendMail

    
    }
    
}


