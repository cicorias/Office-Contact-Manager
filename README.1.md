# Simple Office Mail Addin that Creates a Contact


## Git Clone


## Generate Certificates
Site has to run under SSL (https) so there's a script that will generage a key pair.

```
./makePKI.sh
```

This generaates 3 files:
1. private.pem  (private key)
2. public.pem (public certificate)
3. public.cer (same as above - exact copy)

The .cer file can be used to import directly in Windows certificate management into the trusted authority.

Or use PowerShell

```
Import-Certificate "./public.cer" -CertStoreLocation "Cert:\CurrentUser\root"
```


## Ensure you have App setup in AAD
### Register as a Web App / Web API
record the ClientID and ClientSecret

Review this about ensuring OAuth 2.0 Implict grant flow is allowed:

https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp 

```json
  "keyCredentials": [],
  "knownClientApplications": [],
  "logoutUrl": null,
  "oauth2AllowImplicitFlow": true,
  "oauth2AllowUrlPathMatching": false,
  "oauth2Permissions": [
```


## Tips on Graph and O365 Preview API
Look here for latest: https://graph.microsoft.com/beta/$metadata 


## Run site Using Visual Studio
There is a project file for use in VS 2015 along with the NodeJS tools for Visual Studio.