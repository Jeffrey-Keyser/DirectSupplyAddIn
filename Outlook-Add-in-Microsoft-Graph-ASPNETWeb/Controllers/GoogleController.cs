using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.Graph;
using Newtonsoft.Json;
using OutlookAddinMicrosoftGraphASPNET.Helpers;
using OutlookAddinMicrosoftGraphASPNET.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.UI;

namespace OutlookAddinMicrosoftGraphASPNET.Controllers
{
    public class GoogleController : Controller
    {

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/drive-dotnet-quickstart.json
        static string[] Scopes = { DriveService.Scope.DriveFile };
        static string ApplicationName = "Drive API .NET Quickstart";

        // GET: Google
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Finds a folder with name 'foldername' in Google Drive
        /// </summary>
        /// <returns> Google Drive File Object if found </returns>
        public Google.Apis.Drive.v3.Data.File findFolder(DriveService service, string foldername)
        {
            FilesResource.ListRequest request;

            request = service.Files.List();
            request.Q = $"mimeType = 'application/vnd.google-apps.folder' and name = '{foldername}'";

            var result = request.Execute();
            Google.Apis.Drive.v3.Data.File file;

            // Should only be one, but Google Drive allows for multiple files with same name.
            // Maybe need better checking
            if (result.Files.Count() == 1)
                file = result.Files.FirstOrDefault();
            else
                file = null;


            return file;
        }

        /// <summary>
        /// Creates a folder in Google Drive
        /// </summary>
        /// <param name="parentId">Id of parent folder if subfolder is desired. If null creates in Home directory </param>
        /// <returns> Id of folder created </returns>
        public string createFolderGoogleDrive(DriveService service, string filename, string parentId)
        {
            FilesResource.CreateRequest createRequest;


            Google.Apis.Drive.v3.Data.File body = new Google.Apis.Drive.v3.Data.File();

            body.Name = filename;
            body.MimeType = "application/vnd.google-apps.folder";
            body.Parents = new List<string>();

            if (parentId != null)
            { body.Parents.Add(parentId); }

            createRequest = service.Files.Create(body);

            createRequest.Fields = "id";
            var result = createRequest.Execute();

            return result.Id;
        }

        [System.Web.Http.HttpPost]

        /// <summary>
        /// Saves attachments on Outlook email to Google Drive
        /// </summary>
        /// <returns> Id of folder containing attachments </returns>
        public async Task<string> saveAttachmentGoogleDrive([FromBody] SaveAttachmentRequest request)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                //string credPath = "token.json";
                var credPath = System.Web.HttpContext.Current.Server.MapPath("/App_Data/MyGoogleStorage");


                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }


            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });



            FilesResource.CreateMediaUpload uploadRequest;

            string childId;
            string parentId;
            // Google allows folders with same name.
            // If the Outlook Attachments already exists, don't create another one
            var folder = findFolder(service, "Outlook Attachments");
            parentId = folder.Id;
            if (parentId == null)
            {
                parentId = createFolderGoogleDrive(service, "Outlook Attachments", null);
                childId = createFolderGoogleDrive(service, request.subject, parentId);
            }
            else
            {
                childId = createFolderGoogleDrive(service, request.subject, folder.Id);
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

                    Google.Apis.Drive.v3.Data.File fileMetaData = new Google.Apis.Drive.v3.Data.File();
                    fileMetaData.Name = attachment.Name;
                    fileMetaData.Parents = new List<string>();

                    fileMetaData.Parents.Add(childId); // child folder Id

                    // For files, build a stream directly from ContentBytes
                    if (attachment.Size < (4 * 1024 * 1024))
                    {
                        MemoryStream stream;

                        using (stream = new MemoryStream(Convert.FromBase64String(attachment.ContentBytes)))
                        {
                            uploadRequest = service.Files.Create(
                               fileMetaData, stream, Format.GetMimeType(attachment.Name));
                            uploadRequest.Fields = "id";
                            uploadRequest.Upload();
                        }
                        try
                        {
                            var file = uploadRequest.ResponseBody;
                            Console.WriteLine("File ID: " + file.Id);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        i++;
                    }
                    else
                    {
                        // TODO: Implement functionality to support > 4 MB files
                    }
                }


            }



            return parentId;
        }

        /* Authorize current the same as saveAttachmentGoogleDrive */
        [System.Web.Http.HttpPost]
        public async Task<string> Authorize([FromBody] SaveAttachmentRequest request)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                //string credPath = "token.json";
                var credPath = System.Web.HttpContext.Current.Server.MapPath("/App_Data/MyGoogleStorage");

                
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            
            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            FilesResource.CreateMediaUpload uploadRequest;

            string childId;
            string parentId;
            // Google allows folders with same name.
            // If the Outlook Attachments already exists, don't create another one
            var folder = findFolder(service, "Outlook Attachments");
            parentId = folder.Id;
            if (parentId == null)
            {
                parentId = createFolderGoogleDrive(service, "Outlook Attachments", null);
                childId = createFolderGoogleDrive(service, request.subject, parentId);
            }
            else
            {
                childId = createFolderGoogleDrive(service, request.subject, folder.Id);
            }
            /* Start get attachment */

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

                    Google.Apis.Drive.v3.Data.File fileMetaData = new Google.Apis.Drive.v3.Data.File();
                    fileMetaData.Name = attachment.Name;
                    fileMetaData.Parents = new List<string>();

                    fileMetaData.Parents.Add(childId); // child folder Id

                    // For files, build a stream directly from ContentBytes
                    if (attachment.Size < (4 * 1024 * 1024))
                    {
                        MemoryStream stream;
                        
                        using (stream = new MemoryStream(Convert.FromBase64String(attachment.ContentBytes)))
                        {
                            uploadRequest = service.Files.Create(
                               fileMetaData, stream, Format.GetMimeType(attachment.Name));
                            uploadRequest.Fields = "id";
                            uploadRequest.Upload();
                        }
                        try
                        {
                            var file = uploadRequest.ResponseBody;
                            Console.WriteLine("File ID: " + file.Id);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        i++;
                    }
                    else
                    {
                        // TODO: Implement functionality to support > 4 MB files
                    }
                }


            }


            /* End get attachment */

            return parentId;
        }

    }
}