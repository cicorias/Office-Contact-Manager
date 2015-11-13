
function getContactAddRequest(givenName, surname, fileAs) {
    // Return a GetItem operation request for the subject of the specified item. 
    var result =
    "<soap:Envelope"
    + "    xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\""
    + "    xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\""
    + "    xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\""
    + "    xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">"
    + "     <soap:Body>"
    + "       <CreateItem xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" >"
    + "         <SavedItemFolderId>"
    + "           <t:DistinguishedFolderId Id=\"contacts\"/>"
    + "         </SavedItemFolderId>"
    + "         <Items>"
    + "           <t:Contact>"
    + "             <t:FileAs>" + fileAs + "</t:FileAs>"
    + "             <t:GivenName>" + givenName + "</t:GivenName>"
    + "             <t:CompanyName>Blue Yonder Airlines</t:CompanyName>"
    + "             <t:EmailAddresses>"
    + "               <t:Entry Key=\"EmailAddress1\">tplate@example.com</t:Entry>"
    + "             </t:EmailAddresses>"
    + "             <t:PhysicalAddresses>"
    + "               <t:Entry Key=\"Business\">"
    + "                 <t:Street>1234 56th Ave</t:Street>"
    + "                 <t:City>La Habra</t:City>"
    + "                 <t:State>CA</t:State>"
    + "                 <t:CountryOrRegion>USA</t:CountryOrRegion>"
    + "               </t:Entry>"
    + "             </t:PhysicalAddresses>"
    + "             <t:PhoneNumbers>"
    + "               <t:Entry Key=\"BusinessPhone\">4255550199</t:Entry>"
    + "             </t:PhoneNumbers>"
    + "             <t:JobTitle>Manager</t:JobTitle>"
    + "             <t:Surname>" + surname + "</t:Surname>"
    + "           </t:Contact>"
    + "         </Items>"
    + "       </CreateItem>"
    + "     </soap:Body>"
    + "   </soap:Envelope>";

    return result;
}





function sendRequest() {

    var givenName = $('#givenName').val();
    var surname = $('#surname').val();
    var fileAs = $('#fileAs').val();

    // Create a local variable that contains the mailbox.
    var mailbox = Office.context.mailbox;

    var soapMessage = getContactAddRequest(givenName, surname, fileAs);
    console.log(soapMessage);

    mailbox.makeEwsRequestAsync(soapMessage, soapCallback);
}

function soapCallback(asyncResult) {
    var result = asyncResult.value;
    console.log(asyncResult.value);
    var context = asyncResult.context;



    // Process the returned response here.
}


