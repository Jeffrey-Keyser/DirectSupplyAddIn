using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OutlookAddinMicrosoftGraphASPNET.Models
{
    public class SaveAttachmentRequest
    {
        public string[] filenames { get; set; }
        public string[] attachmentIds { get; set; }
        public string messageId { get; set; }

        public string outlookToken { get; set; }

        public string outlookRestUrl { get; set; }

        // Don't necessarily need this, not in IsValid()
        public string subject { get; set; }

        public bool IsValid()
        {
            return attachmentIds != null && filenames != null &&
                !string.IsNullOrEmpty(messageId) &&
                !string.IsNullOrEmpty(outlookToken) &&
                !string.IsNullOrEmpty(outlookRestUrl);
        }

    }
}