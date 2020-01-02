

using System;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace DocxEmailToWordPress
{
    static class Program
    {
       // <summary>
       // The main entry point for the application.
       // </summary>

       // Docx Email to WordPress website, this program will download emails with D.E.T ST positions
       // and scrap the document for the relevant position information and then repost it to our company website via the
       // word press API
       // Copy-left (Copy left) 2019  John Thompson

        // This program is free software: you can redistribute it and/or modify
        // it under the terms of the GNU General Public License as published by
        // the Free Software Foundation, either version 3 of the License, or
        // (at your option) any later version.

        // This program is distributed in the hope that it will be useful,
        // but WITHOUT ANY WARRANTY; without even the implied warranty of
        // MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
        // GNU General Public License for more details.

        // You should have received a copy of the GNU General Public License
        // along with this program.  If not, see <https://www.gnu.org/licenses/>


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
