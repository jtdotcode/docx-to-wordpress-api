using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenPop.Mime;
using OpenPop.Pop3;
using OpenPop.Common;
using System.IO;

namespace DocxEmailToWordPress
{
    class EmailDownloader
    {
        static String errorTo = "***REMOVED***";
        static String errorFrom = "***REMOVED***";
        static String smtpHost = "***REMOVED***";
        static String tssAddress = "***REMOVED***";
        static String testAddress = "***REMOVED***";

        WordPressApi wordPressApi = new WordPressApi();
        GetWordHtml getWordHtml = new GetWordHtml();
        SendEmail sendEmail = new SendEmail(errorTo, errorFrom, smtpHost);
        String fileExtension = ".docx";

        Int64 EpochLastSent { set; get; }

        // email settings 
        public String hostname = "pop.gmail.com";
        public Int32 port = 995;
        public bool SSL = true;
        private String username = "***REMOVED***";
        private String password = "***REMOVED***";
       

        public EmailDownloader()
        {
            EpochLastSent = (int)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;

            Console.WriteLine("in constuctor");
        }

        public bool TestConnection()
        {
            using (Pop3Client client = new Pop3Client())
            {

                try
                {
                    client.Connect(hostname, port, SSL);
                    client.Authenticate(username, password, AuthenticationMethod.UsernameAndPassword);



                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);


                }

                bool connected = client.Connected;


                if (connected.Equals(true))
                {
                    client.Disconnect();
                    return true;

                }
                else
                {
                    return false;
                }


            }

            

        }

        public void DownloadAttachments()
        {
            using (Pop3Client client = new Pop3Client())
            {
                
                client.Connect(hostname, port, SSL);
                client.Authenticate(username, password, AuthenticationMethod.UsernameAndPassword);
                if (client.Connected)
                {

                    int messageCount = client.GetMessageCount();
                    List<Message> allMessages = new List<Message>(messageCount);
                    for (int i = messageCount; i > 0; i--)
                    {
                        if (!(client.GetMessage(i).Headers.From.Address != tssAddress || client.GetMessage(i).Headers.From.Address != testAddress))
                        {
                            client.DeleteMessage(i);
                            Console.Write("Deleted" + client.GetMessage(i).Headers.From.Address);

                        } else
                        {
                            allMessages.Add(client.GetMessage(i));
                        }
                        

                    }
                    foreach (Message message in allMessages)
                    {
                        Console.Write(message.Headers.From.Address);
                        var attachments = message.FindAllAttachments();
                        foreach (var attachment in attachments)
                        {
                            Int64 msgAttachmentFileSize = attachment.Body.Length;

                            
                            attachment.Save(new System.IO.FileInfo(System.IO.Path.Combine("c:\\emails", attachment.FileName)));

                            Int64 localFileSize = new System.IO.FileInfo(System.IO.Path.Combine("c:\\emails", attachment.FileName)).Length;


                            if (localFileSize == msgAttachmentFileSize)
                            {

                                Console.WriteLine("Local file is: " + localFileSize);
                                Console.WriteLine("Attachment size is: " + msgAttachmentFileSize);
                                var e = Path.GetExtension("c:\\emails\\" + attachment.FileName);
                                var posted = false;

                                if (e == fileExtension)
                                {
                                    var htmldata = getWordHtml.ReadWordDocument("c:\\emails\\" + attachment.FileName);

                                    posted = true;
                                    // = wordPressApi.PostData(htmldata, "test");
                                }
                                else
                                {
                                    try
                                    {
                                        File.Delete("c:\\emails\\" + attachment.FileName);
                                        Console.WriteLine("Deleted " + attachment.FileName);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("Error unable to delete " + attachment.FileName);
                                        Console.WriteLine(ex);
                                    }
                                    
                                }



                                if (posted)
                                {
                                    Console.WriteLine("Successfully Posted");
                                    String successSubject = $"Message from + {message.Headers.From.Address} Post Success";
                               //     sendEmail.Send(successSubject, htmldata);
                                }
                                else
                                {
                                    String messageFrom = "message from:" + message.Headers.From.Address;
                                    String messageSubject = message.Headers.Subject;
                                    String errorBody = $"This message is from {messageFrom}";
                                    String errorSubject = $"Something went wrong {messageFrom} + {messageSubject}";
                                    
                                    Console.WriteLine("Something Went Wrong");
                                  //  sendEmail.Send(errorSubject, errorBody);

                                }






                            }
                            else
                            {
                                Console.WriteLine("Attachment Mismatch");
                                Console.WriteLine("Attachment " + attachment.FileName + " file size should be: " + msgAttachmentFileSize + " Bytes");
                                Console.WriteLine("Disk File is " + localFileSize);
                                Console.WriteLine("Trying to download again");
                                DownloadAttachments();

                            }







                        }


                    }
                }
            }



        } // end Download Attachments


        public void CheckForNewMsg()
        {



            int freqP = 60 * 60;


            while (true)
            {
                Int64 epochCurrentTime = (int)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;

                long timeElapsed = epochCurrentTime - EpochLastSent;


                Console.WriteLine("this is the current time: " + epochCurrentTime);

                Console.WriteLine("this is the lastSent: " + EpochLastSent);

                Console.WriteLine("this many secs have Elapsed: " + timeElapsed);

                if (timeElapsed >= freqP)
                {

                    // send 



                    Console.WriteLine("sending");

                    EpochLastSent = (int)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;



                } // end if


            } //end while



        } // end CheckForNewMsg 


    }
}
