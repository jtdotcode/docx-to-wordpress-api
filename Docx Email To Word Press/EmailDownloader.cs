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
        

        WordPressApi wordPressApi = new WordPressApi();
        GetWordHtml getWordHtml = new GetWordHtml();
        SendEmail sendEmail = new SendEmail(errorTo, errorFrom, smtpHost);
        String fileExtension = ".docx";
        String tmpFolderPath = "c:\\emails\\";

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
                        if (client.GetMessage(i).Headers.From.Address == tssAddress)
                        {
                            allMessages.Add(client.GetMessage(i));

                        } else
                        {
                            var subject = client.GetMessage(i).Headers.Subject;
                            var from = client.GetMessage(i).Headers.From;
                            client.DeleteMessage(i);
                            Console.Write("Deleted" + subject + "From " + from);
                            
                        }
                        

                    }
                    foreach (Message message in allMessages)
                    {
                        Console.Write(message.Headers.From.Address);
                        var attachments = message.FindAllAttachments();
                        foreach (var attachment in attachments)
                        {
                            Int64 msgAttachmentFileSize = attachment.Body.Length;

                            var filePath = tmpFolderPath + attachment.FileName;

                            attachment.Save(new System.IO.FileInfo(System.IO.Path.Combine(tmpFolderPath, attachment.FileName)));

                            Int64 localFileSize = new System.IO.FileInfo(System.IO.Path.Combine(tmpFolderPath, attachment.FileName)).Length;


                            if (localFileSize == msgAttachmentFileSize)
                            {

                                Console.WriteLine("Local file is: " + localFileSize);
                                Console.WriteLine("Attachment size is: " + msgAttachmentFileSize);
                                var e = Path.GetExtension(filePath);
                                var posted = false;

                                if (e == fileExtension)
                                {
                                    var htmldata = getWordHtml.ReadWordDocument(filePath);

                                    posted = wordPressApi.PostData(htmldata, getWordHtml.GetTitle());
                                    var from = message.Headers.From.Address;
                                    var subject = message.Headers.Subject;

                                    if (posted)
                                    {
                                        String successSubject = $"Message from {from} {subject} Post Success";
                                        Console.WriteLine(successSubject);
                                        
                                        // sendEmail.Send(successSubject, htmldata);
                                    }
                                    else
                                    {
                                        
                                        String errorBody = $"This message is from {from}";
                                        String errorSubject = $"Something went wrong {subject} {from}";

                                        Console.WriteLine("Something Went Wrong"  + errorSubject);
                                        //  sendEmail.Send(errorSubject, errorBody);

                                    }


                                }
                                else
                                {
                                    try
                                    {
                                        File.Delete(filePath);
                                        Console.WriteLine("Deleted " + filePath);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("Error unable to delete " + filePath);
                                        Console.WriteLine(ex);
                                    }
                                    
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
