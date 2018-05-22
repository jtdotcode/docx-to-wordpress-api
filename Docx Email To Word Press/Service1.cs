using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace DocxEmailToWordPress
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        EmailDownloader emailDownloader;
        public Double IntervalInSecs { get; set; }

        protected override void OnStart(string[] args)
        {
            // Set up a timer to trigger every minute.  
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = IntervalInSecs * 1000; // 60 seconds  
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();

            // instantiate EmailDownloader Class
            emailDownloader = new EmailDownloader();

            emailDownloader.DownloadAttachments();


        }

        protected override void OnStop()
        {

        }

        public void OnTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            //Check Mail Every X secs

         //   emailDownloader.DownloadAttachments();


            // TODO: monitoring activities here.  

            Console.WriteLine("Checking For New Mail");

           // eventLog1.WriteEntry("Monitoring the System", EventLogEntryType.Information, eventId++);
        }

    }
}
