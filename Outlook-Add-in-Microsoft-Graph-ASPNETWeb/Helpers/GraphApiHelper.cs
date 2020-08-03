// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using Microsoft.Ajax.Utilities;
using Microsoft.Graph;
using System;
using System.IO;
using System.Net.Http.Headers;
using System.Threading.Tasks;

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
        // TEST URL
        internal static string TrackEmailChanges = @"https://graph.microsoft.com/v1.0/me/messages";

        // TODO: Changed to test graph API
        internal static string GetWorkbookSearchUrl(string selectedProperties)
        {
            // Construct URL to search OneDrive for Business for Excel workbooks                
            var workbooksSearchRelativeUrl = "(q = '.xlsx')";
            return BaseMSGraphSearchUrl + workbooksSearchRelativeUrl + selectedProperties;
            //return GetFilesUrl;
        }


        // ADDED
        internal static async Task<IUserMessagesCollectionPage> getConversationIdMessages(string accessToken, string conversationId)
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

               var messages = await graphClient.Me.Messages.Request()
                    .GetAsync();

                return messages;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
                return null;
            }

        }

        // ADDED
        internal static async Task<DriveItem> saveAttachmentOneDrive(string accessToken, string filename, Stream fileContent)
        {
            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", accessToken);
                    }));


            string relativeFilePath = "Outlook Attachments/" + filename;

            try
            {
                // This method only supports files 4MB or less
                DriveItem newItem = await graphClient.Me.Drive.Root.ItemWithPath(relativeFilePath)
                    .Content.Request().PutAsync<DriveItem>(fileContent);

                return newItem;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
                return null;
            }

        }




    }


}

