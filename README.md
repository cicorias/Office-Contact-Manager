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



## Run site Using Visual Studio
There is a project file for use in VS 2015 along with the NodeJS tools for Visual Studio.