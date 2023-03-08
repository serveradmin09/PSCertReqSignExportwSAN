# PS Request Certificate PFX with SAN's and KEY - GUI

Windows Certificate Management doesn't let you quickly create and export a certificate with several SAN's. There are some Powershell Modules, which make it easier and faster to get a certificate / a PFX-file, but I didn't had success to add SAN's in any of them. And here began the nightmare.

The only way was to create a request (CSR) containing the SAN's in the INF-file, submit it to the CA, and export it as PFX. Alright, but some hosts are not accepting PFX-Files for import. So you have to split the PFX file to private key and public certificate and upload them seperately.

It costs ages to do all this steps manually with the windows console- so this was my motivation to create this powershell script.

The GUI is created using Windows Forms (.NET Framework)
See the comments in the script for more details.

This script must run on a computer which is part of a domain with a CA server.
Also make sure, you have sufficient permissions.

Improvement ideas for future releases:

-A file-picker for opensssl.exe
  If OpenSSL for Windows is not installed in the setup-default directory, users will be able to manually browse and pick the openssl.exe.
  
-Auto-read certificate-templates from the CA server and list them in the drop-down menu.
  I didn't implement this, because I wanted to decide for myself which templates are shown. Maybe, someone has a better idea for that.
  
-More options for the certificate request INF-part
  For example: KeyLenght, ProviderName or HashAlgorithm TextFields for the user to adjust these settings. Maybe also to input manual OID's.
  
-ErrorLog / detailled error messages
  For easier problem solving if something dooesn't work
  
-Supporting more OS-locales
  At the moment, only English and German Windows OS is supported. Please don't hesitate to help me with other locales.
  
-Virus-scan warnings
  Some anti-virus software solutions are alerting a virus when scanning the file if it is converted to an *.exe. I could solve this by signing my ps1-file.
  

