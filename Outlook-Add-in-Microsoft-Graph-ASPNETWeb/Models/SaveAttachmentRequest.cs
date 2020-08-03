using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OutlookAddinMicrosoftGraphASPNET.Models
{
    public class SaveAttachmentRequest
    {
        public string filename { get; set; }
        public string attachmentId { get; set; }
        public string messageId { get; set; }

        public string outlookToken { get; set; }

        public string outlookRestUrl { get; set; }

        public bool IsValid()
        {
            return attachmentId != null &&
                !string.IsNullOrEmpty(messageId) &&
                !string.IsNullOrEmpty(outlookToken) &&
                !string.IsNullOrEmpty(outlookRestUrl) &&
                !string.IsNullOrEmpty(filename);
        }

    }
}