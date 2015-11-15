# Simple Office Mail Addin that Creates a Contact

NOTE: This readme refers to the project files in the .\src2 directory which is a Visual Studio 2015 Solution.

This example demonstrates:
1. Making a call to an EWS (Exchange Web Services) call direct from JavaScript to create a contact.
2. TBD


## Exchange Web Services from Office.js 
Review the following MSDN content:
* [Call web services from an Outlook add-in](https://msdn.microsoft.com/en-us/library/office/fp160952.aspx)
* [EWS operations that add-ins support](https://msdn.microsoft.com/en-us/library/office/fp160952.aspx#mod_off15_appscope_CallingWebServices_SupportedEWS)
* [CreateItem operation](https://msdn.microsoft.com/en-us/library/office/aa563797.aspx)
* [CreateItem operation (contact)](https://msdn.microsoft.com/en-us/library/office/aa580529.aspx)


### Permissions
This addin as it calls the EWS SOAP endpoint, requires the permission as follows in the manfifest:

```xml
<OfficeApp ...>
...
  <Permissions>ReadWriteMailbox</Permissions>
...

```

Also review: [Authentication and permission considerations for the makeEwsRequestAsync method](https://msdn.microsoft.com/en-us/library/office/fp160952.aspx#mod_off15_appscope_CallingWebServices_AuthAndPerms)

## Git Clone
normal git stuff here...


## Deployment for Development
The project when opened in VS 2015 allows you to deploy via the VS2015 Office Tools for development.  

You can also deploy the manifest directly to a Mailbox for a single user, or deploy organizationally.

When publishing for deployment inside of Visual Studio the project context menu (right click on project) will
provide a 'Publish' option.  In that Wizard the manifest is built.

Ensure that the HTTPS site is operational after deploying.

For example, this is the deployed manifest when running locally

```xml
<?xml version="1.0" encoding="utf-8"?>
<!--Published:70EDFC97-B41D-43C5-B751-7C00AD999804-->
<!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
    xsi:type="MailApp">
  <Id>bab0661e-47c6-457c-9c06-414b0dd9ba70</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>FooBar Inc.</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="ContactManagerClr" />
  <Description DefaultValue="ContactManagerClr" />
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation 
          DefaultValue="https://localhost:44300/AppRead/Home/Home.html" />
        <RequestedHeight>250</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteMailbox</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
</OfficeApp>
```

