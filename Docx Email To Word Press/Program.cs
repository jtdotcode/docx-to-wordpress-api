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
        private static string  fileName = @"C:\temp\test.docx";

        static void Main()
        {
            //ServiceBase[] ServicesToRun;
            //ServicesToRun = new ServiceBase[]
            //{
            //    new Service1()
            //};
            //ServiceBase.Run(ServicesToRun);

            GetWordPlainText getWordPlainText = new GetWordPlainText(fileName);

            getWordPlainText.ReadWordDocument();

        }
    }
}
