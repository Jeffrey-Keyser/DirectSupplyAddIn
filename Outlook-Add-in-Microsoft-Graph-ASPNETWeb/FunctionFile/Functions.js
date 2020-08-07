/* Source file for UI-less buttons found on emails */

Office.initialize = function () {
}

// Helper function to add a status message to
// the info bar.
function statusUpdate(icon, text) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
        type: "informationalMessage",
        icon: icon,
        message: text,
        persistent: false
    });
}

// Adds text into the body of the item, then reports the results
// to the info bar.
function addTextToBody(text, icon, event) {
    Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                statusUpdate(icon, "\"" + text + "\" inserted successfully.");
            }
            else {
                Office.context.mailbox.item.notificationMessages.addAsync("addTextError", {
                    type: "errorMessage",
                    message: "Failed to insert \"" + text + "\": " + asyncResult.error.message
                });
            }
            event.completed();
        });
}

function addDefaultMsgToBody(event) {
    addTextToBody("Inserted by the Add-in Command Demo add-in.", "blue-icon-16", event);
}