using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenPop.Mime;
using OpenPop.Pop3;
using OpenPop.Common;

namespace DocxEmailToWordPress
{
    class EmailDownloader
    {
        static String errorTo = "***REMOVED***";
        static String errorFrom = "***REMOVED***";
        static String smtpHost = "***REMOVED***";
        static String tssAddress = "tss@edumail.vic.edu.au";
        static String testAddress = "***REMOVED***";

        WordPressApi wordPressApi = new WordPressApi();
        GetWordHtml getWordHtml = new GetWordHtml();
        SendEmail sendEmail = new SendEmail(errorTo, errorFrom, smtpHost);


        Int64 EpochLastSent { set; get; }

        // email settings 
        public String hostname = "pop.gmail.com";
        public Int32 port = 993;
        public bool SSL = true;
        private String userName = "***REMOVED***";
        private String passWord = "***REMOVED***";
       

        public EmailDownloader()
        {
            EpochLastSent = (int)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;

            Console.WriteLine("in constuctor");
        }

        public bool TestConnection()
        {
            Pop3Client client = new Pop3Client();

            client.Connect(hostname, port, SSL);

            client.Authenticate(userName, passWord);


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

        public void DownloadAttachments()
        {
            using (OpenPop.Pop3.Pop3Client client = new Pop3Client())
            {
                client.Connect(hostname, port, false);
                client.Authenticate(userName, passWord, AuthenticationMethod.UsernameAndPassword);
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
                    foreach (Message msg in allMessages)
                    {
                        Console.Write(msg.Headers.From.Address);
                        var att = msg.FindAllAttachments();
                        foreach (var ado in att)
                        {
                            Int64 msgAttachmentFileSize = ado.Body.Length;

                            ado.Save(new System.IO.FileInfo(System.IO.Path.Combine("c:\\emails", ado.FileName)));

                            Int64 localFileSize = new System.IO.FileInfo(System.IO.Path.Combine("c:\\emails", ado.FileName)).Length;


                            if (localFileSize == msgAttachmentFileSize)
                            {

                                Console.WriteLine("Local file is: " + localFileSize);
                                Console.WriteLine("Attachment size is: " + msgAttachmentFileSize);
                                var htmldata = getWordHtml.ReadWordDocument("c:\\emails\\" + ado.FileName);
                                var posted = wordPressApi.PostData(htmldata);

                                if (posted)
                                {
                                    Console.WriteLine("Successfully Posted");
                                    String successSubject = $"Message from + {msg.Headers.From.Address} Post Success";
                                    sendEmail.Send(successSubject, htmldata);
                                }
                                else
                                {
                                    String messageFrom = "message from:" + msg.Headers.From.Address;
                                    String messageSubject = msg.Headers.Subject;
                                    String errorBody = $"This message is from {messageFrom}";
                                    String errorSubject = $"Something went wrong {messageFrom} + {messageSubject}";
                                    
                                    Console.WriteLine("Something Went Wrong");
                                    sendEmail.Send(errorSubject, errorBody);

                                }






                            }
                            else
                            {
                                Console.WriteLine("Attachment Mismatch");
                                Console.WriteLine("Attachment " + ado.FileName + " file size should be: " + msgAttachmentFileSize + " Bytes");
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
