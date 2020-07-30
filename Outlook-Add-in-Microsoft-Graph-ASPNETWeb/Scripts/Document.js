// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

"use strict";

let dialog;

Office.initialize = function () {
    $(document).ready(function () {
        app.initialize();


        $("#getOneDriveFilesButton").click(getFileNamesFromGraph); // Dismissed currently
        $("#logoutO365PopupButton").click(logout);
        $("#getEmailInfo").click(getEmailInfo);
        $("#getAttachmentInfo").click(getFileNamesFromGraph);
        $("#deleteCurrentEmail").click(deleteCurrentEmail);
        $("#downloadAttachmentLocal").click(downloadAttachmentsLocally);

    });
};

function getFileNamesFromGraph() {

    $("#instructionsContainer").hide();
    $("#waitContainer").show();

    console.debug("About to call Graph API")

    $.ajax({
        url: "/files/onedrivefiles",
        type: "GET"
    })
    .done(function (result) {
        writeFileNamesToMessage(result)
            .then(function () {
                $("#waitContainer").hide();
                $("#finishedContainer").show();
            })
            .catch(function (error) {
                app.showNotification(error.toString());
            });
    })
        .fail(function (result) {
            app.showNotification("Cannot get data from MS Graph: " + result.toString());
    });
}

function writeFileNamesToMessage(graphData) {

    // Office.Promise is an alias of OfficeExtension.Promise. Only the alias
    // can be used in an Outlook add-in.
    return new Office.Promise(function (resolve, reject) {
        try {
            Office.context.mailbox.item.body.getTypeAsync(
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        app.showNotification(result.error.message);
                    }
                    else {
                        // Successfully got the type of item body.
                        if (result.value === Office.MailboxEnums.BodyType.Html) {

                            // Body is of type HTML.
                            var htmlContent = createHtmlContent(graphData);

                            Office.context.mailbox.item.body.setSelectedDataAsync(
                                htmlContent, { coercionType: Office.CoercionType.Html },
                                function (asyncResult) {
                                    if (asyncResult.status ===
                                        Office.AsyncResultStatus.Failed) {
                                        console.log(asyncResult.error.message);
                                    }
                                    else {
                                        console.log("Successfully set HTML data in item body.");
                                    }
                                });
                        }
                        else {
                            // Body is of type text. 
                            var textContent = createTextContent(graphData);

                            Office.context.mailbox.item.body.setSelectedDataAsync(
                                textContent, { coercionType: Office.CoercionType.Text },
                                function (asyncResult) {
                                    if (asyncResult.status ===
                                        Office.AsyncResultStatus.Failed) {
                                        console.log(asyncResult.error.message);
                                    }
                                    else {
                                        console.log("Successfully set text data in item body.");
                                    }
                                });
                        }
                    }
                });
            resolve();
        }
        catch (error) {
            reject(Error("Unable to add filenames to document. " + error));
        }
    });
}

function createHtmlContent(data) {

    var bodyContent = "<html><head></head><body>";

    for (var i = 0; i < data.length; i++) {
        bodyContent += "<p>" + data[i] + "</p>";
    }
    bodyContent += "</body></html >";

    return bodyContent;
}

function createTextContent(data) {

    var bodyContent = "";
    for (var i = 0; i < data.length; i++) {
        bodyContent += data[i] + "\n";
    }

    return bodyContent;
}

function logout() {

    Office.context.ui.displayDialogAsync('https://localhost:44301/azureadauth/logout',
        { height: 30, width: 30 }, function (result) {           
            dialog = result.value;
            dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processLogoutMessage);
        });
}

function processLogoutMessage(messageFromLogoutDialog) {

    if (messageFromLogoutDialog.message === "success") {
        dialog.close();
        document.location.href = "/home/index";
    }
    else {
        dialog.close();
        app.showNotification("Not able to logout: " + messageFromLogoutDialog.toString());
    }
}


function getItemRestId() {
    if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
        // itemId is already REST-formatted.
        return Office.context.mailbox.item.itemId;
    } else {
        console.debug("Item Converted");
        // Convert to an item ID for API v2.0.
        return Office.context.mailbox.convertToRestId(
            Office.context.mailbox.item.itemId,
            Office.MailboxEnums.RestVersion.v2_0
        );
    }
}

function getAttachmentRestId(attachmentId) {
    if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
        // itemId is already REST-formatted.
        return attachmentId;
    } else {
        console.debug("Attachment Converted");
        // Convert to an item ID for API v2.0.
        return Office.context.mailbox.convertToRestId(
            attachmentId,
            Office.MailboxEnums.RestVersion.v2_0
        );
    }
}


function deleteCurrentEmail() {
    var item = Office.context.mailbox.item;

    console.debug("Button pressed: " + item.itemId);


    // REST call
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === "succeeded") {
            var accessToken = result.value;
            console.debug("accessToken: " + accessToken);

            // Use the access token.
            // execute the function for the Mail API call
           // executeCall(accessToken, args);


            // Delete email call
            var itemId = getAttachmentRestId(item.itemId);

            console.debug("Rest item id: " + itemId);
            // Construct the REST URL to the current item.
            // Details for formatting the URL can be found at
            // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-attachments.
            var getMessageUrl = Office.context.mailbox.restUrl +
                '/v2.0/me/messages/' + itemId;

            console.debug("get request url: DELETE EMAIL " + getMessageUrl);

            $.ajax({
                url: getMessageUrl,
                type: 'DELETE',
                dataType: 'json',
                headers: { 'Authorization': 'Bearer ' + accessToken }
            }).done(function (item) {
                // Message is passed in `item`.
                console.debug("message delete returned: " + JSON.stringify(item))


            }).fail(function (error, textStatus, errorThrown) {
                console.debug("Error after sending ajax call " + textStatus + " Error thrown: " + errorThrown);
                // Handle error.
            });




        } else {
            // Handle the error.
        }
    });

}



function getEmailInfo() {

    // REST call
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === "succeeded") {
            var accessToken = result.value;
            console.debug("accessToken: " + accessToken);

            // Use the access token.
            // execute the function for the Mail API call
            // executeCall(accessToken, args);

            // Get the item's REST ID.
            var itemId = getItemRestId();
            console.debug("Rest item id: " + itemId);
            // Construct the REST URL to the current item.
            // Details for formatting the URL can be found at
            // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
            var getMessageUrl = Office.context.mailbox.restUrl +
                '/v2.0/me/messages/' + itemId;

            console.debug("get request url: ITEM " + getMessageUrl);

            $.ajax({
                url: getMessageUrl,
                dataType: 'json',
                headers: { 'Authorization': 'Bearer ' + accessToken }
            }).done(function (item) {

                // Message is passed in `item`.
                console.debug("Sender info: " + JSON.stringify(item["Sender"]));
                console.debug("Reciever Info: " + JSON.stringify(item["ToRecipients"]));

                console.debug("Conversation id: " + JSON.stringify(item["ConversationId"]));

            }).fail(function (error, textStatus, errorThrown) {
                console.debug("Error after sending ajax call " + textStatus + " Error thrown: " + errorThrown);
                // Handle error.
            });



        } else {
            // Handle the error.
        }
    });



}

function downloadAttachmentsLocally() {

    var item = Office.context.mailbox.item;

    if (item.attachments.length > 0) {
        for (var i = 0; i < item.attachments.length; i++) {
            // TODO: ContentBytes?
            var attachment = item.attachments[i];

            downloadAttachment(attachment.id);
        }
    }

}


function downloadAttachment(attachmentId) {


    // REST call
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === "succeeded") {
            var accessToken = result.value;
            console.debug("accessToken: " + accessToken);

            // Use the access token.
            // execute the function for the Mail API call
            // executeCall(accessToken, args);

            // Get the item's REST ID.
            var itemId = getItemRestId();
            var attachmentRestId = getAttachmentRestId(attachmentId);
            console.debug("accessToken : " + accessToken);
            console.debug("attachmentId : " + attachmentId);
            console.debug("Rest item id: " + itemId);
            // Construct the REST URL to the current item.
            // Details for formatting the URL can be found at
            // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-attachments.
            var getMessageUrl = Office.context.mailbox.restUrl +
                '/v2.0/me/messages/' + itemId + '/attachments/' + attachmentRestId;

            console.debug("get request url: ATTACHMENT CONTENTS " + getMessageUrl);

            $.ajax({
                url: getMessageUrl,
                dataType: 'json',
                headers: { 'Authorization': 'Bearer ' + accessToken }
            }).done(function (item) {

                // Message is passed in `item`.

                var contentBytes = JSON.stringify(item["ContentBytes"]);
                var contentType = JSON.stringify(item["ContentType"]);

                // Get rid of leading / ending double quotes
                contentBytes = contentBytes.replace(/['"]+/g, '');
                contentType = contentType.replace(/['"]+/g, '');

                console.debug("Content Bytes: " + contentBytes + " Type : " + contentType);

                download(contentBytes, item["Name"], contentType);

            }).fail(function (error, textStatus, errorThrown) {
                console.debug("Error after sending ajax call " + textStatus + " Error thrown: " + errorThrown);
                // Handle error.
            });



        } else {
            // Handle the error.
        }
    });


}

// Function to download data to a file
function download(data, filename, type) {

    console.debug("\nEnter Download");

    var readableData = atob(data);

    console.debug("readableData: " + readableData);

    var file = new Blob([readableData], { type: type });
    if (window.navigator.msSaveOrOpenBlob) // IE10+
        window.navigator.msSaveOrOpenBlob(file, filename);
    else { // Other browsers
        var a = document.createElement("a"),
            url = URL.createObjectURL(file);
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        setTimeout(function () {
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
        }, 0);
    }
}
