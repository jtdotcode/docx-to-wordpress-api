using System;
using System.Collections.Generic;
using System.Linq;
using OpenPop.Mime;
using OpenPop.Pop3;
using System.IO;
using RestSharp;

namespace DocxEmailToWordPress
{
    class EmailDownloader
    {

        

        // log4net class log name
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static String smtpSendTo = Properties.Settings.Default.smtpSendTo;
        static String smtpSentFrom = Properties.Settings.Default.smtpSentFrom;


        static String smtpHost = Properties.Settings.Default.smtpHost;
        static String smtpUsername = Properties.Settings.Default.smtpUsername;
        static String smtpPassword = Properties.Settings.Default.smtpPassword;
        static int smtpPort = Properties.Settings.Default.smtpPort;
        static bool smtpSSL = Properties.Settings.Default.smtpSSL;


        static String[] allowedEmailAddress = delimitAllowedEmails(Properties.Settings.Default.allowedEmailAddress);
        

        WordPressApi wordPressApi = new WordPressApi();
        
        SendEmail sendEmail = new SendEmail(smtpHost, smtpSendTo, smtpSentFrom,  smtpPort, smtpUsername, smtpPassword, smtpSSL);
        String fileExtension = ".docx";
        String DocumentFileNameContains = Properties.Settings.Default.DocumentFileNameContains;
        String tmpFolderPath = Path.GetTempPath();
        PostLog log = new PostLog();
        List<PostLog> emailLog = new List<PostLog>();


        // receive email settings 
        public String popHost = Properties.Settings.Default.popHost;
        public Int32 popPort = Properties.Settings.Default.popPort;
        public bool popSSL = Properties.Settings.Default.popSSL;
        private String popUsername = Properties.Settings.Default.popUsername;
        private String popPassword = Properties.Settings.Default.popPassword;
        public Int32 messageLeft = 0;

        public bool TestConnection()
        {
            using (Pop3Client client = new Pop3Client())
            {

                try
                {
                    client.Connect(popHost, popPort, popSSL);
                    client.Authenticate(popUsername, popPassword, AuthenticationMethod.UsernameAndPassword);



                }
                catch (Exception ex)
                {

                    
                    logger.Fatal(ex);

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

        private static string[] delimitAllowedEmails(string allowedEmailsString)
        {

            var allowedEmails = new List<string>();

            if (allowedEmailsString.Contains(','))
            {
                allowedEmails = allowedEmailsString.Split(',').ToList();

                return allowedEmails.ToArray();
            }

            allowedEmails.Add(allowedEmailsString);


            return allowedEmails.ToArray();


        }

        public string checkBccAddress(string s)
        {
            if (s == String.Empty || s == null)
            {
                var emptyTo = "Hidden BCC Address";
                return emptyTo;
            }


            return s;
        }


        public Boolean DownloadAttachments()
        {
            bool noMessages = false;

            using (Pop3Client client = new Pop3Client())
            {
                
                client.Connect(popHost, popPort, popSSL);
                client.Authenticate(popUsername, popPassword, AuthenticationMethod.UsernameAndPassword);

                var messageNum = 0;

                // check if server is connected
                if (client.Connected)
                {
                    
                    // get total message count in the in box
                    int messageCount = client.GetMessageCount();

                    if(messageCount == 0)
                    {
                        logger.Info("There are no new messages to Download, Exiting");
                    }

                    List<Message> allMessages = new List<Message>(messageCount);

                    // count down the total messages
                    for (int i = messageCount; i > 0; i--)
                    {


                        logger.Info("Total Emails to Download are : " + messageCount);
                        logger.Info("Processing " + i + " of " + messageCount);

                        // create log with email details
                        emailLog.Add(new PostLog()
                        {
                            Body = client.GetMessage(i).MessagePart.Body,
                            CurrentDateTime = DateTime.Now,
                            FromAddress = client.GetMessage(i).Headers.From.Address,
                            Subject = client.GetMessage(i).Headers.Subject,
                            MessageCount = messageCount,
                            MessageOf = i,
                            ToAddress = checkBccAddress(client.GetMessage(i).Headers.To.First().Address),
                            TimeRecieved = client.GetMessage(i).Headers.Date,
                            Messages = new List<String>(),
                            Attachments = new Dictionary<String, long>()

                        });




                        // check if the message is from specific sender address, else delete the message
                        if (allowedEmailAddress.Contains(client.GetMessage(i).Headers.From.Address))
                        {
                            // add each message to a List<Message> Array
                            allMessages.Add(client.GetMessage(i));

                            logger.Info("Adding email for processing from " + client.GetMessage(i).Headers.From.Address + " Subject " + client.GetMessage(i).Headers.Subject);
                 

                        } else
                        {
                            
                            // delete the message if not from specific sender

                            var subject = client.GetMessage(i).Headers.Subject;
                            var from = client.GetMessage(i).Headers.From.Address;

                            client.DeleteMessage(i);

                            // add Error message to Messages List Array in PostData
                            
                           emailLog.ElementAt(messageNum).Messages.Add("Email not from " + allowedEmailAddress.ToString() + " Deleting " + subject + "From " + from);

                            logger.Info("Email not from " + allowedEmailAddress.ToString() + " Deleted email subject is: " + subject + "email is from: " + from);

                            break;
                            
                        }


                        


                        //need to fix this NOT going to work.
                        messageNum = messageNum++;

                    }

                    messageNum = 0;
                    var messageLeft = allMessages.Count;

                    // enumerate each message 
                    foreach (Message message in allMessages)
                    {
                        messageLeft--;

                        // Add all attachments for each message a List Array
                        var attachments = message.FindAllAttachments();

                        // enumerate each attachment 
                        foreach (var attachment in attachments)
                        {
                            // get the attachment size
                            Int64 msgAttachmentFileSize = attachment.Body.Length;

                            
                            // set the folder to save the attachments to.
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

                                //if the attachment doesn't match the fileExtension type delete it

                                if (exetension == fileExtension && attachment.FileName.Contains(DocumentFileNameContains))
                                {
                                    String htmldata;

                                    // text from the docx file and return a HTML table
                                    using (GetWordHtml getWordHtml = new GetWordHtml()) {

                                        // add new attachment name to html body
                                        getWordHtml.AttachmentName = attachment.FileName;

                                       
                                        // read the Docx file and return the Schools Table as a HTML table. 
                                        htmldata = getWordHtml.ReadWordDocument(filePath);

                                        IRestResponse responseData = null;

                                        if (htmldata != null)
                                        {

                                            // post HTML table from docx
                                            responseData = wordPressApi.PostData(htmldata, getWordHtml.GetTitle());

                                            // check if successful
                                            posted = responseData.IsSuccessful;

                                            // update List<PostLog> element Post with the Returned Data
                                            emailLog.ElementAt(messageNum).PostStatus = responseData.ResponseStatus.ToString();

                                            logger.Info("File: " + filePath + " Posted Status: " + responseData.ResponseStatus.ToString());


                                        }
                                        else
                                        {

                                            posted = false;
                                            logger.Error("File: " + filePath + " Contains an 0 amount of hours or something went wrong");

                                        }

                                        

                                        

                                        

                                        
                                        // if the post is unsuccessful record response data for REST API to logger4net log
                                        if (!posted)
                                        {
                                            if (responseData != null) {

                                                logger.Debug("Error Exception: " + responseData.ErrorException);
                                                logger.Debug("Post Error StatusCode: " + responseData.StatusCode);
                                                logger.Debug("Post Error Message: " + responseData.ErrorMessage);

                                            }
                                            
                                        }

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
                                            logger.Fatal(ex);
                                            
                                        }
                                        


                                    }
                                    else
                                    {
                                        

                                        // log if unable to post
                                        emailLog.ElementAt(messageNum).Messages.Add($"Something went wrong with the post {subject} {from} - {currentTime} ");
                                        logger.Info($"Something went wrong with the post! Email: {subject} {from} Attachment: {filePath} ");
                                        

                                        if (!attachment.FileName.Contains(DocumentFileNameContains)){

                                            logger.Info($"Document File Name doesnt Contain: {DocumentFileNameContains} FileName is: {filePath} ");
                                        }

                                        // trying to remove file
                                        try
                                        {
                                            if (File.Exists(filePath)){

                                                File.Delete(filePath);

                                            }
                                            
                                        }
                                        catch (Exception ex)
                                        {
                                            logger.Fatal(ex);

                                        }


                                    }


                                }
                                else
                                {
                                    try
                                    {
                                        // deleting the Attachment

                                        File.Delete(filePath);

                                        

                                        // log deleted message
                                        emailLog.ElementAt(messageNum).Messages.Add("Non Docx Attachment Found, Deleted " + filePath + " " + DateTime.Now);
                                        logger.Info("Non Docx Attachment Found, Deleted " + filePath + " ");
                                    }
                                    catch (Exception ex)
                                    {
                                        

                                        // log failed to delete 
                                        emailLog.ElementAt(messageNum).Messages.Add("Error unable to delete " + filePath + ex);
                                        logger.Fatal("Error unable to delete " + filePath + " Exception: " + ex);
                                    }
                                    
                                }
                                
                            }
                            else
                            {
                                // attachment size mismatch
                                emailLog.ElementAt(messageNum).Messages.Add("Attachment Mismatch" + attachment.FileName + " file size should be: " + msgAttachmentFileSize + " Bytes");

                                logger.Info("Attachment Mismatch " + attachment.FileName + " file size should be: " + msgAttachmentFileSize + " Bytes");


                                // Trying to download again
                                logger.Info("Trying to download attachments again");

                                DownloadAttachments();

                               

                                return false;

                            }

                            
                        }


                        Message lastItem = allMessages.Last();

                        // if the there is no more messages 
                        if (message.Equals(lastItem))
                        {

                            noMessages = true;
                            sendEmail.Send(emailLog);

                            messageLeft = 0;



                        }


                        messageNum++;

                    } // end for-each loop for messages



                } // end check if server is connected
            }

            return noMessages;

        } // end Download Attachments


        


    }
}
