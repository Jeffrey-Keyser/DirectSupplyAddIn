using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OutlookAddinMicrosoftGraphASPNET.Models
{
    public class UploadSessionResponse
    {
        public string odatacontext { get; set; }
        public DateTime expirationDateTime { get; set; }
        public string[] nextExpectedRanges { get; set; }
        public string uploadUrl { get; set; }
    }
}