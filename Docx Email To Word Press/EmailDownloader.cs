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
        static String smtpSendTo = "***REMOVED***";
        static String smtpSentFrom = "***REMOVED***";
        static String smtpHost = "***REMOVED***";
        static int smtpPort = 587;
        static String tssAddress = "***REMOVED***";
        

        WordPressApi wordPressApi = new WordPressApi();
        
        SendEmail sendEmail = new SendEmail(smtpHost, smtpSendTo, smtpSentFrom,  smtpPort);
        String fileExtension = ".docx";
        String tmpFolderPath = "c:\\emails\\";
        PostLog log = new PostLog();
        List<PostLog> emailLog = new List<PostLog>();


        // email settings 
        public String hostname = "pop.gmail.com";
        public Int32 port = 995;
        public bool SSL = true;
        private String username = "***REMOVED***";
        private String password = "***REMOVED***";
        public Int32 messageLeft = 0;

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

        public Boolean DownloadAttachments()
        {
            bool noMessages = false;

            using (Pop3Client client = new Pop3Client())
            {
                
                client.Connect(hostname, port, SSL);
                client.Authenticate(username, password, AuthenticationMethod.UsernameAndPassword);

                var messageNum = 0;

                // check if server is connected
                if (client.Connected)
                {
                    
                    // get total message count in the inbox
                    int messageCount = client.GetMessageCount();

                    List<Message> allMessages = new List<Message>(messageCount);

                    // count down the total messages
                    for (int i = messageCount; i > 0; i--)
                    {
                        // log message count
                        messageLeft = i;

                        // check if the message is from specific sender address, else delete the message
                        if (client.GetMessage(i).Headers.From.Address == tssAddress)
                        {
                            // add each message to a List<Message> Array
                            allMessages.Add(client.GetMessage(i));

                            // create log with email details
                            emailLog.Add(new PostLog() { Body = client.GetMessage(i).MessagePart.Body,
                                CurrentDateTime = DateTime.Now, FromAddress = client.GetMessage(i).Headers.From.Address,
                                Subject = client.GetMessage(i).Headers.Subject, MessageCount = messageCount, MessageOf = i,
                                ToAddress = client.GetMessage(i).Headers.To.First().Address, TimeRecieved = client.GetMessage(i).Headers.Date, Messages = new List<String>(), Attachments = new Dictionary<String, long>()

                            });

                        } else
                        {
                            // delete the message if not from specific sender

                            var subject = client.GetMessage(i).Headers.Subject;
                            var from = client.GetMessage(i).Headers.From.Address;

                            client.DeleteMessage(i);

                            // add Errormessage to Messages List Array in PostData
                            emailLog.ElementAt(messageNum).Messages.Add( "Email not from " + tssAddress + " Deleting " + subject + "From " + from );
                            Console.Write("Deleted" + subject + "From " + from);
                            
                        }


                        //need to fix this NOT going to work.
                        messageNum = messageNum++;

                    }

                    messageNum = 0;

                    // enumerate each message 
                    foreach (Message message in allMessages)
                    {
                        
                        // Add all attachments for each message a List Array
                        var attachments = message.FindAllAttachments();

                        // enumerate each attachment 
                        foreach (var attachment in attachments)
                        {
                            // get the attachment size
                            Int64 msgAttachmentFileSize = attachment.Body.Length;

                            
                            // set the folder to save the attactments to.
                            var filePath = @tmpFolderPath + attachment.FileName;

                            // save the attachment to the computer
                            attachment.Save(new System.IO.FileInfo(System.IO.Path.Combine(tmpFolderPath, attachment.FileName)));

                            // get the local attachment file size
                            Int64 localFileSize = new System.IO.FileInfo(System.IO.Path.Combine(tmpFolderPath, attachment.FileName)).Length;

                            // log file name and file size
                            emailLog.ElementAt(messageNum).Attachments.Add(attachment.FileName, msgAttachmentFileSize);

                            if (localFileSize == msgAttachmentFileSize)
                            {

                                // get extension for each attachment
                                var exetension = Path.GetExtension(filePath);

                                var posted = false;

                                //if the attachemnt doesnt match the fileExtension type delete it

                                if (exetension == fileExtension)
                                {
                                    String htmldata;

                                    // text from the docx file and return a html table
                                    using (GetWordHtml getWordHtml = new GetWordHtml()) {
                                        htmldata = getWordHtml.ReadWordDocument(filePath);

                                        // post html table from docx
                                        var responseData = wordPressApi.PostData(htmldata, getWordHtml.GetTitle());

                                        // check if successful
                                        posted = responseData.IsSuccessful;

                                        // update List<PostLog> element Post with the Returned Data
                                        emailLog.ElementAt(messageNum).PostStatus = responseData.ResponseStatus.ToString();

                                    } 

                                    
                                    var from = message.Headers.From.Address;
                                    var subject = message.Headers.Subject;
                                    var currentTime = DateTime.Now.ToShortDateString();

                                    if (posted)
                                    {
                                        // record time posted with from and subject add to the Messages <List>
                                        emailLog.ElementAt(messageNum).Messages.Add($"Message from {from} {subject} Post Success - {currentTime} ");

                                        // update Posted htmldate for log
                                        emailLog.ElementAt(messageNum).PostedHtml = htmldata;

                                        // set posted for Email Log.
                                        emailLog.ElementAt(messageNum).Posted = posted;

                                        // remove file from temp folder

                                        try
                                        {
                                            File.Delete(filePath);
                                        }
                                        catch (Exception ex)
                                        {

                                            Console.Write(ex.ToString());
                                        }
                                        


                                    }
                                    else
                                    {
                                        

                                        // log if unable to post
                                        emailLog.ElementAt(messageNum).Messages.Add($"Something went wrong with the post {subject} {from} - {currentTime} ");

                                      

                                    }


                                }
                                else
                                {
                                    try
                                    {
                                        // deleting the Attachment

                                        File.Delete(filePath);

                                        

                                        // log deleted message
                                        emailLog.ElementAt(messageNum).Messages.Add("No Docx Attachment Found, Deleted " + filePath + DateTime.Now);
                                    }
                                    catch (Exception ex)
                                    {
                                        

                                        // log failed to delete 
                                        emailLog.ElementAt(messageNum).Messages.Add("Error unable to delete " + filePath + ex);
                                    }
                                    
                                }
                                
                            }
                            else
                            {
                                // attachment size mismatch
                                emailLog.ElementAt(messageNum).Messages.Add("Attachment Mismatch" + attachment.FileName + " file size should be: " + msgAttachmentFileSize + " Bytes");
                                
                                // Trying to download again
                                DownloadAttachments();

                                return false;

                            }


                            // if the there is no more messages 
                            if (messageLeft == 1)
                            {

                                noMessages = true;
                                sendEmail.Send(emailLog);

                                messageLeft = 0;

                            } else if( messageLeft == 0)
                            {

                                noMessages = false;
                            }




                        }

                        messageNum++;

                    } // end foreach loop for messages



                } // end check if server is connected
            }

            return noMessages;

        } // end Download Attachments


        


    }
}
