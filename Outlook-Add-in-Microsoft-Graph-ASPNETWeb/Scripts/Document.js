﻿// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

"use strict";

let dialog;

Office.initialize = function () {
    $(document).ready(function () {
        app.initialize();


      //  $("#getOneDriveFilesButton").click(getFileNamesFromGraph); // Dismissed currently
        $("#logoutO365PopupButton").click(logout);
        $("#getConversationWithId").click(getConversationWithId);
        $("#saveAttachmentsOneDrive").click(saveAttachmentsOneDrive);
        $("#deleteCurrentEmail").click(deleteCurrentEmail);
     //   $("#downloadAttachmentLocal").click(downloadAttachmentsLocally);

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

/*
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
} */

function createHtmlContent(data, elementId) {

    for (var i = 0; i < data.length; i++) {
        var file = document.createElement('p');
        file.className = 'ms-font-l ms-fontColor-themePrimary indentFromPaneEdge centeredText';
        file.innerHTML = data[i];
        document.getElementById(elementId).appendChild(file);
    }

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



// Get messages in same conversation as current email
// Does a graph API call
function getConversationWithId() {
    
    // REST call
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === "succeeded") {
            var accessToken = result.value;
            // Use the access token.
            // execute the function for the Mail API call
            // executeCall(accessToken, args);

            // Get the item's REST ID.
            var itemId = getItemRestId();

            console.debug("accessToken : " + accessToken);
            console.debug("Rest item id: " + itemId);
            // Construct the REST URL to the current item.
            // Details for formatting the URL can be found at
            // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-attachments.
            var getMessageUrl = Office.context.mailbox.restUrl +
                '/v2.0/me/messages/' + itemId;

            console.debug("get request url: ATTACHMENT CONTENTS " + getMessageUrl);

            // Get conversation id
            $.ajax({
                url: getMessageUrl,
                dataType: 'json',
                headers: { 'Authorization': 'Bearer ' + accessToken }
            }).done(function (item) {

                var conversationId = JSON.stringify(item["ConversationId"]);


                console.debug("Conversation ID : " + conversationId);

                // Get rid of leading / ending double quotes
                conversationId = conversationId.replace(/['"]+/g, '');

                // Controller call does a GRAPH API call.
                $.ajax({
                    url: "/files/conversationmessages",
                    type: "GET",
                    data: { convoId: conversationId },
                })
                    .done(function (result) {

                        console.debug("Success: " + result);

                    })
                    .fail(function (result) {
                        app.showNotification("Cannot get data from MS Graph: " + result.toString());
                    });



            }).fail(function (error, textStatus, errorThrown) {
                console.debug("Error after sending ajax call " + textStatus + " Error thrown: " + errorThrown);
                // Handle error.
            });



        } else {
            // Handle the error.
        }
    });

}



function saveAttachmentsOneDrive() {


    var item = Office.context.mailbox.item;
    var attachmentIds = [];
    var filenames = [];
    // For each attachment
    for (var i = 0; i < item.attachments.length; i++) {

        var attachmentRestId = getAttachmentRestId(item.attachments[i].id);

        filenames.push(item.attachments[i].name);
        attachmentIds.push(attachmentRestId);
    }

    // REST call to get token for ContentBytess
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === "succeeded") {

            var accessToken = result.value;

            console.debug("accessToken : " + accessToken);
            console.debug("subject : " + item.subject);

            var saveAttachmentRequest = {
                filenames: filenames,
                attachmentIds: attachmentIds,
                messageId: getItemRestId(),
                outlookToken: accessToken,
                outlookRestUrl: Office.context.mailbox.restUrl,
                subject: item.subject,
            }

            // Controller call to save attachment to oneDrive
            $.ajax({
                url: "/files/saveattachmentonedrive",
                type: "POST",
                data: JSON.stringify(saveAttachmentRequest),
                contentType: "application/json; charset=utf-8"
            })
                .done(function (result) {

                    console.debug("Successful save: attachmentUrl " + result);

                    var attachmentUrls = result;

                    // Delete the local attachments
                    $.ajax({
                        url: "/files/deleteemailattachments",
                        type: "POST",
                        data: { attachmentIds: attachmentIds, emailId: getItemRestId(), attachmentUrls: result },
                    })
                        .done(function (result) {

                            console.debug("Success: " + result);

                            // Embed link to OneDrive Location
                            embedAttachmentLinks(attachmentUrls, getItemRestId(), accessToken);

                            createHtmlContent(filenames, 'finishedContainer');

                            // Display result for 5 seconds
                            $("#instructionsContainer").hide();
                            $("#finishedContainer").show();

                            setTimeout(showCommands, 5000);

                        })
                        .fail(function (result) {
                            app.showNotification("Cannot get data from MS Graph: " + result.toString());
                        });


                })
                .fail(function (result) {
                    app.showNotification("Cannot get data from MS Graph: " + result.toString());
                });



        } else {
            // Handle the error.
        }
    });

}

// Callback function for timeout
function showCommands() {
    $("#finishedContainer").hide();
    $("#instructionsContainer").show();
}

function embedAttachmentLinks(attachmentsLocation, emailId, accessToken) {

    $.ajax({
        url: "/email/addattachmentstobody",
        type: "GET",
        data: { attachmentsLocation: attachmentsLocation, emailId: emailId },
    })
        .done(function (result) {

            console.debug("Success embed : " + result.toString());


            var getMessageUrl = Office.context.mailbox.restUrl +
                '/v2.0/me/messages/' + emailId;

            var message = {
                Body: {
                    "ContentType": '1',
                    "Content": result.toString(),
                }
            };

            $.ajax({
                url: getMessageUrl,
                contentType: 'application/json',
                type: 'PATCH',
                headers: { 'Authorization': 'Bearer ' + accessToken },
                data: JSON.stringify(message),
            }).done(function (item) {

                console.debug("Success, email is updated");

            })
                .fail(function (xhr, textStatus, errorThrown) {
                    app.showNotification("Error: Couldn't update email with new body: " + textStatus);
                    var jsonResponse = JSON.parse(xhr.responseText);
                    console.debug(jsonResponse);
                });
        })
}
