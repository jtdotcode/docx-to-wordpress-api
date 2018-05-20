﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxEmailToWordPress
{
    class PostLog
    {

        public String FromAddress { get; set; }
        public String ToAddress { get; set; }
        public String Subject { get; set; }
        public List<String> Message { get; set; }
        public byte[] Body { get; set; }
        public DateTime CurrentDateTime { get; set; }
        public Int32 MessageCount { get; set; }
        public Boolean Attachment { get; set; }
        public Dictionary<String, Int64> Attachments { get; set; }
        public String ErrorMessage { get; set; }
    }
}