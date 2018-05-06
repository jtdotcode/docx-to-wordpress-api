using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Docx_Email_To_Word_Press
{
    class WpBlogPosts
    {
        public String Title { get; set; }
        public String Message { get; set; }
        public DateTime PublishedOn { get; set; }
        public String Comments { get; set; }
        public String Link { get; set; }

    }
}
