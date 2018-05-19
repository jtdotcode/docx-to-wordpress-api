using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace DocxEmailToWordPress
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// 
        

        static void Main()
        {
            //ServiceBase[] ServicesToRun;
            //ServicesToRun = new ServiceBase[]
            //{
            //    new Service1()
            //};
            //ServiceBase.Run(ServicesToRun);


            WordPressApi wordPressApi = new WordPressApi();
            GetWordHtml getWordHtml = new GetWordHtml();
            EmailDownloader emailDownloader = new EmailDownloader();

            // wordPressApi.PostData(getWordHtml.ReadWordDocument(@"c:\\temp\\test.docx"), getWordHtml.GetTitle());

            // getWordHtml.ReadWordDocument(@"c:\\temp\\test.docx");

            // jsonData.GetHtmlData(dic);

            if (emailDownloader.TestConnection())
            {
                Console.WriteLine("Can connect");
                emailDownloader.DownloadAttachments();

            }
            else
            {
                Console.WriteLine("Something Went Wrong");
            }
            


        }
    }
}
