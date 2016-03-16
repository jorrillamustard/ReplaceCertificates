# ReplaceCertificates

REPLACECERTS.EXE uses openSSL to create and replace Certificates on Fidelis Endpoint. There is also the option to define your own pre-created certificates.

Format: 
replacecerts.exe -n <\New Public Cert Name> <\New Private Cert Name>

replacecerts.exe -d <\New Public Cert Name> <\New Private Cert Name>

switches:

-n : Creates new certificates using openSSL and changes the certificates in Site Server and Work Manager configurations

-d : Takes pre-created certificates as commands and changes the configurations of site server and work manager (Does not create new certificates)

Download Release at:

https://github.com/jorrillamustard/ReplaceCertificates/releases/latest
