// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using Microsoft.Ajax.Utilities;
using Microsoft.Graph;
using System;
using System.IO;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Mvc;
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


    }
}

