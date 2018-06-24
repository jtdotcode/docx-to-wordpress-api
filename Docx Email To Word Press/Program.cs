

using System;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace DocxEmailToWordPress
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// 

        // log4net class log name
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
            static void Main(string[] args)
            {

           
            EmailDownloader emailDownloader = new EmailDownloader();
            WordPressApi wordPressApi = new WordPressApi();

            if (emailDownloader.TestConnection().Equals(true) && wordPressApi.testApi().Equals(true))
            {
                logger.Info("Accounts Test Successful");
                emailDownloader.DownloadAttachments();
                
            }
            else
            {
                
                if (emailDownloader.TestConnection().Equals(false))
                {
                    logger.Error("Something Went Wrong, Unable to connect to Pop Email Account");
                    logger.Error("Exiting");
                }
                if (wordPressApi.testApi().Equals(false))
                {
                    logger.Error("Something Went Wrong, Unable to connect to WP API Account");
                    logger.Error("Exiting");
                }


                

            }

            

        }
    }
}
