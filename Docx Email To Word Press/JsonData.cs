using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxEmailToWordPress
{
    class JsonData
    {


        public String Title { get; set; }
        public String Status { get; set; }
        public String Content { get; set; }
        public String Excerpt { get; set; }
        public String[] Categories { get; set; }




        public String BuildHtmlTable(Dictionary<String, Double> dictionary)
        {
            StringBuilder sbHours = new StringBuilder();
            StringBuilder sbSchools = new StringBuilder();
            Double totalHours = 0.0;

            foreach (var item in dictionary)
            {
                var x = item.Value;
                totalHours = x + totalHours;

                try
                {
                    sbSchools.Append(item.Key.ToString());
                    sbSchools.Append("<br />");
                    
                    sbHours.Append(item.Value.ToString());
                    sbHours.Append("<br />");

                }
                catch (Exception)
                {

                    throw;
                }
                

                    
            }

            String schools =  sbSchools.ToString();
            String hours = sbHours.ToString();
            String totalHoursString = totalHours.ToString();



           String htmltable = $"<table border=\\\"1\\\" cellspacing=\\\"0\\\" cellpadding=\\\"0\\\" width=\\\"638\\\"><tr><td width=\\\"508\\\"><h3><strong>School Name<\\/strong><\\/h1><\\/td><td width=\\\"130\\\"><h3 align=\\\"center\\\"><strong>Hours<br />(Per Week)<\\/strong><\\/h3><\\/td><\\/tr><tr><td width=\\\"508\\\" valign=\\\"top\\\"><p>{schools}<br /><\\/td><td width =\\\"130\\\" valign=\\\"top\\\"><p align=\\\"center\\\">{hours}<br /><\\/td><\\/tr><tr><td width=\\\"508\\\"><p align=\\\"right\\\"><strong>Total Per Week<\\/strong><\\/p><\\/td><td width=\\\"130\\\"><p align=\\\"center\\\"><strong>{totalHoursString}<\\/strong><strong> <\\/strong><\\/p><\\/td><\\/tr><\\/table>";

            Console.WriteLine(htmltable);

        return htmltable;
    }
        


    }
}
