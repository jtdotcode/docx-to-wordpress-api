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
        // WpAPI wpAPI = new WpAPI();
        DocxToString docxToString = new DocxToString();


        Int64 epochLastSent { set; get; }

        // email settings 
        public String hostname = "pop.gmail.com";
        public Int32 port = 993;
        public bool SSL = true;
        private String userName = "test";
        private String passWord = "test";


        public EmailDownloader()
        {
            epochLastSent = (int)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
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
                        allMessages.Add(client.GetMessage(i));

                    }
                    foreach (Message msg in allMessages)
                    {
                        Console.Write(msg.Headers.From.Address);
                        var att = msg.FindAllAttachments();
                        foreach (var ado in att)
                        {
                            Int64 msgAttachmentFileSize = ado.Body.Length;

                            ado.Save(new System.IO.FileInfo(System.IO.Path.Combine("c:\\temp", ado.FileName)));

                            Int64 localFileSize = new System.IO.FileInfo(System.IO.Path.Combine("c:\\temp", ado.FileName)).Length;


                            if (localFileSize == msgAttachmentFileSize)
                            {

                                Console.WriteLine("Local file is: " + localFileSize);
                                Console.WriteLine("Attachment size is: " + msgAttachmentFileSize);

                                //var body = audioToText.AsyncRecognize("c:\\temp\\" + ado.FileName);
                                //sendMail.Send(xmlsettings.ToEmail, msg.Headers.From.Address, "message from:" + msg.Headers.From.Address, body.ToString());


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

                long timeElapsed = epochCurrentTime - epochLastSent;


                Console.WriteLine("this is the current time: " + epochCurrentTime);

                Console.WriteLine("this is the lastSent: " + epochLastSent);

                Console.WriteLine("this many secs have Elapsed: " + timeElapsed);

                if (timeElapsed >= freqP)
                {

                    // send 



                    Console.WriteLine("sending");

                    epochLastSent = (int)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;



                } // end if


            } //end while



        } // end CheckForNewMsg 


    }
}
