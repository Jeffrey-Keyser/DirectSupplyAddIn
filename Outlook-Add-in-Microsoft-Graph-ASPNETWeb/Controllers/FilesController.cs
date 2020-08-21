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
using System.Text;

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
        /// Saves all attachments found on an email to OneDrive
        /// </summary>
        /// <returns> Urls of the saved attachments </returns>
        [System.Web.Http.HttpPost]
        public async Task<string> SaveAttachmentOneDrive([FromBody]SaveAttachmentRequest request)
        {

            if (request == null || !request.IsValid() || request.attachmentIds.Length == 0)
            {
                return null;
            }

            string attachmentsUrl = null; 

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


                        // Get access token from SQL database
                        var token = Data.GetUserSessionToken(Settings.GetUserAuthStateId(ControllerContext.HttpContext), Settings.AzureADAuthority);

                        // TODO: Check if the file already exists 

                        attachmentsUrl =  await GraphApiHelper.saveAttachmentOneDrive(token.AccessToken, Format.MakeFileNameValid(request.filenames[i]), fileStream, Format.MakeFileNameValid(request.subject));

                        // Format 
                        string delete = "/" + request.filenames[i];
                        attachmentsUrl = attachmentsUrl.Replace(delete, "");

                        i++;
                    }
                    else
                    {
                        // Functionality to support > 4 MB files
                        // See https://docs.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0
                        var token = Data.GetUserSessionToken(Settings.GetUserAuthStateId(ControllerContext.HttpContext), Settings.AzureADAuthority);

                        DriveItem folder = await GraphApiHelper.searchFileOneDrive(token.AccessToken, "Outlook Attachments");


                        var url = "https://graph.microsoft.com/v1.0" + $"/me/drive/items/{folder.Id}:/{attachment.Name}:/createUploadSession";
                        var uploadReq = new HttpRequestMessage(HttpMethod.Post, url);


                        // Headers
                        uploadReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.AccessToken);

                        // Send request
                        var sessionResponse = client.SendAsync(uploadReq).Result.Content.ReadAsStringAsync().Result;


                        var uploadSession = JsonConvert.DeserializeObject<UploadSessionResponse>(sessionResponse);

                        Upload upload = new Upload();

                        HttpResponseMessage response = upload.UploadFileBySession(uploadSession.uploadUrl, Convert.FromBase64String(attachment.ContentBytes));


                        return folder.WebUrl;
                    }
                }  

                return attachmentsUrl;

            }

        }


    }
}
