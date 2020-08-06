// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Net.Http;
using System.Web.Http;
using System.Net;
using OutlookAddinMicrosoftGraphASPNET.Helpers;
using OutlookAddinMicrosoftGraphASPNET.Models;
using System;
using Microsoft.Ajax.Utilities;
using Microsoft.Graph;
using System.Web.Http.Results;
using System.IdentityModel;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.IO;
using System.Runtime.CompilerServices;

namespace OutlookAddinMicrosoftGraphASPNET.Controllers
{
    public class FilesController : Controller
    {
        /// <summary>
        /// Recursively searches OneDrive for Business.
        /// </summary>
        /// <returns>The names of the first three workbooks in OneDrive for Business.</returns>
        public async Task<JsonResult> OneDriveFiles()
        {

            // Get access token
            var token = Data.GetUserSessionToken(Settings.GetUserAuthStateId(ControllerContext.HttpContext), Settings.AzureADAuthority);

            // Get all the Excel files in OneDrive for Business by using the Microsoft Graph API. Select only properties needed.
            var fullWorkbooksSearchUrl = GraphApiHelper.GetWorkbookSearchUrl("?$select=name,id&top=3");
            var filesResult = await ODataHelper.GetItems<ExcelWorkbook>(fullWorkbooksSearchUrl, token.AccessToken);

            List<string> fileNames = new List<string>();
            foreach(ExcelWorkbook workbook in filesResult)
            {
                fileNames.Add(workbook.Name);
            }
            return Json(fileNames, JsonRequestBehavior.AllowGet); 
        }

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
        /// Saves all attachments found on an email to OneDrive
        /// </summary>
        /// <returns> Urls of the saved attachments </returns>
        [System.Web.Http.HttpPost]
        public async Task<String[]> SaveAttachmentOneDrive([FromBody]SaveAttachmentRequest request)
        {
            string[] attachmentsUrl = new string[request.attachmentIds.Length]; 

            if (request == null || !request.IsValid())
            {
                return null;
            }

            using (var client = new HttpClient())
            {
                // Get content bytes
                string baseAttachmentUri = request.outlookRestUrl;
                if (!baseAttachmentUri.EndsWith("/"))
                    baseAttachmentUri += "/";
                baseAttachmentUri += "v2.0/me/messages/" + request.messageId + "/attachments/";

                var i = 0;
                foreach (string attachmentId in request.attachmentIds)
                {

                    var getAttachmentReq = new HttpRequestMessage(HttpMethod.Get, baseAttachmentUri + attachmentId);

                    // Headers
                    getAttachmentReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", request.outlookToken);
                    getAttachmentReq.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    var result = await client.SendAsync(getAttachmentReq);
                    string json = await result.Content.ReadAsStringAsync();
                    OutlookAttachment attachment = JsonConvert.DeserializeObject<OutlookAttachment>(json);

                    // For files, build a stream directly from ContentBytes
                    if (attachment.Size < (4 * 1024 * 1024))
                    {
                        MemoryStream fileStream = new MemoryStream(Convert.FromBase64String(attachment.ContentBytes));


                        // Get access token
                        var token = Data.GetUserSessionToken(Settings.GetUserAuthStateId(ControllerContext.HttpContext), Settings.AzureADAuthority);

                        attachmentsUrl[i] =  await GraphApiHelper.saveAttachmentOneDrive(token.AccessToken, MakeFileNameValid(request.filenames[i]), fileStream, request.subject);
                        
                        i++;
                    }
                    else
                    {
                        // TODO: Implement functionality to support > 4 MB files

                        return null;
                    }
                }

                return attachmentsUrl;

            }

        }

        // Helper function that formats filenames for OneDrive upload
        private string MakeFileNameValid(string originalFileName)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            return string.Join("_", originalFileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries)).TrimEnd('.');
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

    }
}
