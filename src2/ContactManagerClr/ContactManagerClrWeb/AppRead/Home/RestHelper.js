function sendRequestViaServer() {

    var emailToken = app.session.callbacktoken;  //$('#callbacktoken').html();
    var identityToken = app.session.identitytoken; //$('#identitytoken').html();
    var ewsUrl = app.session.ewsUrl;


    var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
    app.session.itemid = item.itemId;


    var data = {
        EmailToken: emailToken,
        IdentityToken: identityToken,
        ItemId: app.session.itemid,
        EWSUrl: ewsUrl
    }

    $.ajax({
        type: "POST",
        url: "/api/Ews",
        data: data,
        success: postComplete,
        dataType: "application/json"
    }).fail(function (e) {
        console.error('failed to call');
    });


}


function postComplete() {
    console.log('done');

}