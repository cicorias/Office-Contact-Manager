/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            //displayItemDetails();

            console.log('starting callbacks');
            Office.context.mailbox.getCallbackTokenAsync(tokenCallBack);
            Office.context.mailbox.getUserIdentityTokenAsync(userIdentityCallback);

            console.log('done callbacks');

        });
    };

    function tokenCallBack(asyncResult) {
        console.log('token call back start');
        //https://msdn.microsoft.com/EN-US/library/office/jj984589.aspx
        if (asyncResult.status === "failed") {
            $('#callbacktoken').html(error.message);
            console.log('failed on tokencallback');
        }
        else {
            console.log(asyncResult.value);
            app.session.callbacktoken = asyncResult.value;
            $('#callbacktoken').html(asyncResult.value);
        }
    }

    function userIdentityCallback(asyncResult) {
        if (asyncResult.status === "failed") {
            $('#identitytoken').html(error.message);
        }
        else {
            console.log("Bearer " + asyncResult.value);
            app.session.identitytoken = asyncResult.value;
            $('#identitytoken').html(asyncResult.value);
        }
    }


    // Displays the "Subject" and "From" fields, based on the current mail item
    function displayItemDetails() {
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        $('#subject').text(item.subject);

        var from;
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            from = Office.cast.item.toMessageRead(item).from;
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            from = Office.cast.item.toAppointmentRead(item).organizer;
        }

        if (from) {
            $('#from').text(from.displayName);
            $('#from').click(function () {
                app.showNotification(from.displayName, from.emailAddress);
            });
        }
    }
})();