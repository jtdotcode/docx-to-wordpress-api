

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

            if (emailDownloader.TestConnection())
            {
                logger.Info("Email Account Test Successful");
                emailDownloader.DownloadAttachments();

            }
            else
            {
                logger.Info("Something Went Wrong, Unable to connect to Pop Email Account");
                
            }



            



        }
    }
}
