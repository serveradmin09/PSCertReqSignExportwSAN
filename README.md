# PSCertReqSignExportwSAN

Windows Certificate Management doesn't let you quickly create and export a certificate with several SAN's.
There are some Powershell Modules, which make it easier and faster to get a certificate in a pfx-file, but I didn't has success to add SAN's in any of them.
And here begins the nightmare. 
The only way was to create a request (CSR) containing the SAN's in the INF-file, submit it to the CA, and export it as PFX.
Especially if you are PKI-admin of a big company and you have to manage hundreds of certificates.
Some systems are not accepting PFX-Files.
So you have to split the PFX file into private key and public certificate and upload them seperate.
It costs ages to do all this steps manually - so this was my motivation to create this script.

The GUI is created using Windows Forms (.NET Framework)
See the comments in the script for more details.

Improvement ideas for future releases:

-A text-box for setting an own passphrase.
  For now, the passphrase for the private key is the hostname, written backwards. My idea is to use a own passphrase if entered it in the textbox.

-A file-picker for opensssl.exe
  If OpenSSL for Windows is not installed in the setup-default directory, the user is able to manually browse and pick the openssl.exe.
  
-Auto-read certificate-templates from the CA server and list them in the drop-down menu.
  I didn't implement this, because I wanted to decide for myself which templates are shown. Maybe, someone has a better idea for that.
  
-More options for teh certificate request INF-part
  For example: KeyLenght, ProviderName or HashAlgorithm TextFields for the user to adjust these settings. Maybe also to input manual OID's.
  
-ErrorLog / detailed error messages
  For easier problem solving if something dooesn't work
  
-Supporting more OS-locales
  At the moment, only English and German Windows OS is supported. Please don't hesitate to help me with other locales.
  
-Virus-scan warnings
  Some anti-virus software solutions are alerting a virus when scanning the file if it is converted to an *.exe. I could solve this by signing my ps1-file.
  

