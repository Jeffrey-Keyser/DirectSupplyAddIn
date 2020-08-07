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
        /// Deletes all attachments on current email
        /// </summary>
        /// <returns> String array of deleted attachments </returns>
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
                img = getBetween(body, "<img ", ">");

                // Replace with hyperlink to attachment
                string delete = "<img " + img + ">";
                hyperlink = "<a href=\"" + attachmentsLocation + "\"> View Attachments on OneDrive </a>";

                if (oneLinkAvailable && img != "")
                {
                    body = body.Replace(delete, hyperlink);
                    oneLinkAvailable = false;
                }
                else
                    body = body.Replace(delete, "");
            } while (img != "");

            string pageBreak = "<div style=\"font - family:Calibri,Arial,Helvetica,sans - serif; font - size:12pt; color: rgb(0, 0, 0)\"><br></div>";
            // If no embedded attachments, add a link at the beginng of the email
            if (oneLinkAvailable)
            {
                string oldBody = getBetween(body, "<body dir=\"ltr\">", "</body>");
                string newBody = hyperlink + pageBreak + oldBody;

                oldBody = "<body dir=\"ltr\">" + oldBody + "</body>";
                newBody = "<body dir=\"ltr\">" + newBody + "</body>";

                body = body.Replace(oldBody, newBody);

            }


            return body;

        }


        // Helper function for addAttachmentsToBody
        // Gets i
        public static string getBetween(string strSource, string strStart, string strEnd)
        {
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                int Start, End;
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }

            return "";
        }

    }
}