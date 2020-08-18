// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using Microsoft.Ajax.Utilities;
using Microsoft.Graph;
using Microsoft.OData.Edm.EdmToClrConversion;
using Microsoft.Office365.OutlookServices;
using OutlookAddinMicrosoftGraphASPNET.Controllers;
using OutlookAddinMicrosoftGraphASPNET.Models;
using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Remoting.Messaging;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Windows.Forms;

namespace OutlookAddinMicrosoftGraphASPNET.Helpers
{
    /// <summary>
    /// Provides methods for Microsoft Graph-specific endpoints.
    /// </summary>
    internal static class GraphApiHelper
    {
        // Microsoft Graph-related base URLs
        internal static string GetFilesUrl = @"https://graph.microsoft.com/v1.0/me/drive/root/sharedWithMe";
        internal static string BaseMSGraphSearchUrl = @"https://graph.microsoft.com/v1.0/me/drive/root/microsoft.graph.search";
        // internal static string BaseItemsUrl = @"https://graph.microsoft.com/1/me/drive/items/";

        internal static string GetWorkbookSearchUrl(string selectedProperties)
        {
            // Construct URL to search OneDrive for Business for Excel workbooks                
            var workbooksSearchRelativeUrl = "(q = '.xlsx')";
            return BaseMSGraphSearchUrl + workbooksSearchRelativeUrl + selectedProperties;
            //return GetFilesUrl;
        }

        /// <summary>
        /// Retrieves all messages with the specified conversation id
        /// </summary>
        /// <returns> All messages in the conversation </returns>
        internal static async Task<IUserMessagesCollectionPage> getConversationIdMessages(string accessToken, string conversationId)
        {
            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", accessToken);
                    }));

            string filterString = $"conversationId eq '{conversationId}'";

            try
            {

               var messages = await graphClient.Me.Messages.Request()
                    .Filter(filterString)
                    .GetAsync();

                var myEmail = await graphClient.Me.Request()
                    .GetAsync();

                // Find the latest reply from another source.
                // Prevents deletion of email chain
                int latestReplyIndex = 0;
                int currIndex = 0;
                foreach ( var email in messages.CurrentPage)
                {
                    // If not equal to my email, set as last index
                    if (email.Sender.EmailAddress.Address != myEmail.Mail)
                    {
                        latestReplyIndex = currIndex;
                    }
                    currIndex++;
                }


                // Delete all older emails in convo
                // Should be sorted already so just delete in order
                int i = 0;
                foreach ( var email in messages.CurrentPage)
                {
                    if (i < latestReplyIndex)
                    {
                        // If earlier email has attachments, prompt user to ensure 
                        if (email.Attachments != null)
                        {
                            // TODO: Handle
                            MessageBox.Show("Attachments will be deleted..");
                        }

                        await graphClient.Me.Messages[email.Id]
                            .Request()
                            .DeleteAsync();
                    }
                    i++;
                }

                return messages;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
                return null;
            }

        }

        /// <summary>
        /// Saves attachment to OneDrive under the parent 'Outlook Attachments' and under child 'Email subject' 
        /// </summary>
        /// <returns> Attachment's url location </returns>
        internal static async Task<String> saveAttachmentOneDrive(string accessToken, string filename, Stream fileContent, string subject)
        {
            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", accessToken);
                    }));

            string relativeFilePath = "Outlook Attachments/" + subject + "/" + filename;

            try
            {
                // This method only supports files 4MB or less
                DriveItem newItem = await graphClient.Me.Drive.Root.ItemWithPath(relativeFilePath)
                    .Content.Request().PutAsync<DriveItem>(fileContent);

                // Embed url in the email.
                // newItem.WebUrl

                return newItem.WebUrl;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
                return null;
            }

        }

        /// <summary>
        /// Deletes attachments located on the selected email
        /// </summary>
        /// <returns> Returns true if attachments were deleted, false otherwise </returns>
        public static async Task<bool> deleteEmailAttachments(string accessToken, string[] attachmentIds, string emailId, string[] attachmentUrls)
        {

            if (attachmentIds == null)
            {
                return false;
            }


            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", accessToken);
                    }));


            foreach (string attachmentId in attachmentIds)
            {
                try
                {
                    // Delete from email
                    await graphClient.Me.Messages[emailId]
                        .Attachments[attachmentId]
                        .Request()
                        .DeleteAsync();


                    // TODO: Attachment preview deletion? Gave ErrorNotFound

                }
                catch (Exception ex)
                {
                    System.Diagnostics.Trace.WriteLine(ex.ToString());
                    return false;
                }
            }

            return true;

        }

        /// <returns> Body of current email </returns>
        internal static async Task<string> getMessageBody(string accessToken, string emailId)
        {
            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", accessToken);
                    }));


            try
            {
                // Just get the email body for processing later
                var email = await graphClient.Me.Messages[emailId]
                    .Request()
                    .GetAsync();

                // In HTML
                return email.Body.Content;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
                return null;
            }

        }


        internal static async Task<Microsoft.Graph.IMailFolderMessagesCollectionPage> getMailFolderMessages(string folderId, string accessToken, string requestUri, string callbackToken)
        {

            var graphClient = new GraphServiceClient(
               new DelegateAuthenticationProvider(
                   async (requestMessage) =>
                   {
                       requestMessage.Headers.Authorization =
                           new AuthenticationHeaderValue("Bearer", accessToken);
                   }));

            Microsoft.Graph.MailFolder mailFolder;

            EmailController control = new EmailController();
            // TESTING. For now just do inbox folder.
            // Later do based on folderId
            try
            {

                mailFolder = await graphClient.Me.MailFolders
                                        .Inbox
                                        .Request()
                                        .GetAsync();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
                return null;
            }



            try
            {
                int count = 0;

                Microsoft.Graph.IMailFolderMessagesCollectionPage messages;

                // First call to get the first 10 messages
                messages = await graphClient.Me.MailFolders[mailFolder.Id]
                                    .Messages
                                    .Request()
                                    .GetAsync();


                foreach( var message in messages.CurrentPage)
                {
                    if ((message.HasAttachments != null && message.HasAttachments == true) || (Format.getBetween(message.Body.Content.ToString(), "<img", ">") != "") )
                    {
                        var attachments = await graphClient.Me.Messages[message.Id]
                                            .Attachments
                                            .Request()
                                            .GetAsync();

                        string[] attachmentIds = new string[attachments.CurrentPage.Count];
                        int index = 0;
                        string[] attachmentUrls = null;
                        string[] attachmentNames = new string[attachments.CurrentPage.Count];
                        string attachmentLocation = null;
                        foreach (Microsoft.Graph.FileAttachment attachment in attachments.CurrentPage)
                        {
                            attachmentIds[index] = attachment.Id;
                            attachmentNames[index++] = attachment.Name;

                            string attachmentContent = Convert.ToBase64String(attachment.ContentBytes);

                            // For files, build a stream directly from ContentBytes
                            if (attachment.Size < (4 * 1024 * 1024))
                            {
                                MemoryStream fileStream = new MemoryStream(Convert.FromBase64String(attachmentContent));

                                // OneDrive Save
                               // attachmentLocation = await saveAttachmentOneDrive(accessToken, attachment.Name, fileStream, message.Subject);



                            }

                        }

                        // Google Drive save
                        GoogleController google = new GoogleController();
                        SaveAttachmentRequest newRequest = new SaveAttachmentRequest()
                        {
                            filenames = attachmentNames,
                            attachmentIds = attachmentIds,
                            messageId = message.Id,
                            outlookRestUrl = requestUri,
                            outlookToken = callbackToken,
                            subject = message.Subject
                        };

                        // Google Drive save
                        attachmentLocation = await google.saveAttachmentGoogleDrive(newRequest);

                        attachmentLocation = "https://drive.google.com/drive/u/1/folders/" + attachmentLocation;

                        // Delete
                        await deleteEmailAttachments(accessToken, attachmentIds, message.Id, attachmentUrls);

                        // Patch new body to email
                        await control.updateEmailBody(requestUri, attachmentLocation, message.Id, accessToken, callbackToken);
                        /*
                        using (var client = new HttpClient())
                        {
                            // Embed link
                            var body = await control.addAttachmentsToBody(attachmentLocation, message.Id, accessToken);

                            // Post updated body
                            var method = new HttpMethod("PATCH");

                            requestUri += "/v2.0/me/messages/" + message.Id;


                            var iContent = new StringContent("{ 'Body' : {" +
                                    " 'ContentType': '1', 'Content': '" + body.ToString() + "'} }", System.Text.Encoding.UTF8, "application/json");

                            string myContent = await iContent.ReadAsStringAsync();

                            var request = new HttpRequestMessage(method, requestUri)
                            {
                                Content = iContent
                            };

                            // Headers
                            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", callbackToken);
                            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                            var result = await client.SendAsync(request);
                        }
                        */
                    }
                }

                // Page size is 10
                count += 10;

                // Iterate through all
                while (mailFolder.TotalItemCount > count)
                {

                    messages = await graphClient.Me.MailFolders[mailFolder.Id]
                                    .Messages
                                    .Request()
                                    .Skip(count)
                                    .GetAsync();

                    // Page size is 10
                    count += 10;
                }

                return messages;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
                return null;
            }

        }


    }
}

