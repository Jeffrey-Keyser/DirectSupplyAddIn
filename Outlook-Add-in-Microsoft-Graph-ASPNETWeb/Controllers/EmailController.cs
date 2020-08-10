using OutlookAddinMicrosoftGraphASPNET.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace OutlookAddinMicrosoftGraphASPNET.Controllers
{
    public class EmailController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }


        /// <summary>
        /// Gets emails with conversation id 
        /// </summary>
        /// <returns>Emails with the specific conversation id.</returns>
        public async Task<JsonResult> ConversationMessages(string convoId)
        {

            // Get access token
            var token = Data.GetUserSessionToken(Settings.GetUserAuthStateId(ControllerContext.HttpContext), Settings.AzureADAuthority);

            var messages = await GraphApiHelper.getConversationIdMessages(token.AccessToken, convoId);

            return Json(messages, JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// Deletes all attachments on current email
        /// </summary>
        /// <returns> String array of deleted attachments </returns>
        [System.Web.Http.HttpPost]
        public async Task<dynamic> deleteEmailAttachments(string[] attachmentIds, string emailId, string[] attachmentUrls)
        {
            // Get access token
            var token = Data.GetUserSessionToken(Settings.GetUserAuthStateId(ControllerContext.HttpContext), Settings.AzureADAuthority);

            var attachments = await GraphApiHelper.deleteEmailAttachments(token.AccessToken, attachmentIds, emailId, attachmentUrls);

            return Json(attachments, JsonRequestBehavior.AllowGet);

        }


        /// <summary>
        /// Adds hyperlink to OneDrive in email's body if attachments present.
        /// </summary>
        /// <returns> Email's new body with hyperlink </returns>
        public async Task<string> addAttachmentsToBody(string attachmentsLocation, string emailId)
        {
            // Get access token
            var token = Data.GetUserSessionToken(Settings.GetUserAuthStateId(ControllerContext.HttpContext), Settings.AzureADAuthority);

            string body = await GraphApiHelper.getMessageBody(token.AccessToken, emailId);

            // For embedded images, remove <img > before adding links
            string img;
            string hyperlink;
            bool oneLinkAvailable = true;
            do
            {
                img = Format.getBetween(body, "<img ", ">");

                // Replace with hyperlink to attachment
                string delete = "<img " + img + ">";
                hyperlink = "<a href=\"" + attachmentsLocation + "\"> View Attachments on OneDrive </a>";

                // Replace first <img > with hyperlink
                if (oneLinkAvailable && img != "")
                {
                    body = body.Replace(delete, hyperlink);
                    oneLinkAvailable = false;
                }
                else
                    body = body.Replace(delete, "");
            } while (img != "");


            string pageBreak = "<div style=\"font - family:Calibri,Arial,Helvetica,sans - serif; font - size:12pt; color: rgb(0, 0, 0)\"><br></div>";
            
            // If no embedded attachments, add a link at the beginning of the email
            if (oneLinkAvailable)
            {
                string oldBody = Format.getBetween(body, "<body dir=\"ltr\">", "</body>");
                string newBody = hyperlink + pageBreak + oldBody;

                oldBody = "<body dir=\"ltr\">" + oldBody + "</body>";
                newBody = "<body dir=\"ltr\">" + newBody + "</body>";

                body = body.Replace(oldBody, newBody);

            }
            return body;
        }
    }
}