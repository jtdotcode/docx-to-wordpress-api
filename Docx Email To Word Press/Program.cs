using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

[assembly: log4net.Config.XmlConfigurator(Watch=true)]

namespace DocxEmailToWordPress
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// 

        // log4net class log name
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
            static void Main(string[] args)
            {
            


            WordPressApi wordPressApi = new WordPressApi();
            GetWordHtml getWordHtml = new GetWordHtml();
           
            //EmailDownloader emailDownloader = new EmailDownloader();

            // wordPressApi.PostData(getWordHtml.ReadWordDocument(@"c:\\temp\\test.docx"), getWordHtml.GetTitle());

            // getWordHtml.ReadWordDocument(@"c:\\temp\\test.docx");

            // jsonData.GetHtmlData(dic);

            //if (emailDownloader.TestConnection())
            //{
            //    Console.WriteLine("Can connect");
            //    emailDownloader.DownloadAttachments();

            //}
            //else
            //{
            //    Console.WriteLine("Something Went Wrong");
            //}

            EmailDownloader emailDownloader = new EmailDownloader();

            emailDownloader.DownloadAttachments();


            Console.Read();



        }
    }
}
