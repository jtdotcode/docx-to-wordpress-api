using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace DocxEmailToWordPress
{
    class SendEmail
    {

        public string username = "***REMOVED***";
        public string Password = "***REMOVED***";
        public string _host { set; get; }
        public string _errorTo { set; get; }
        public string _errorFrom { set; get; }

        public SendEmail(String errorTo, String errorFrom, String host)
        {
            _errorTo = errorTo;
            _errorFrom = errorFrom;
            _host = host;

        }
      

            public Boolean Send(string subject, string body)
            {
                SmtpClient smtpClient = new SmtpClient();
                NetworkCredential basicCredential =
                    new NetworkCredential(username, Password);
                MailMessage message = new MailMessage();
                MailAddress fromAddress = new MailAddress(_errorFrom);

                smtpClient.Host = _host;
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = basicCredential;

                message.From = fromAddress;
                message.Subject = subject;
                //Set IsBodyHtml to true means you can send HTML email.
                message.IsBodyHtml = false;
                message.Body = body;
                message.To.Add(_errorTo);
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



        } // end sendMail
    
}


