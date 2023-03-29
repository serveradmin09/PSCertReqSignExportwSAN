# ---------------------------------------------------------------------------------------------#
# This script creates a certificate with subject saternative names (DNS), using a CSR          #
# imports it into the Windows certificate storage (my certificates) on your issuing CA-server  #
# exports the public certificate as CER-File,                                                  #
# the public certificate including private key as PFX-File and                                 #
# the private key as KEY-File to your destination directory .                                  #
# ---------------------------------------------------------------------------------------------#

# checking if OpenSSL for Windows is installed #
if (!(Test-Path "C:\Program Files\OpenSSL-Win64\bin\openssl.exe") -or (Test-Path "C:\Program Files (x86)\OpenSSL-Win32\bin\openssl.exe"))
    {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show('OpenSSL for Windows was not found in the setup-default directory. Please install OpenSSL first!','Warning')
    } 
    else {}


# creating list of CA's - maximum is 5 #
    $cadump = certutil -dump | select-string -Pattern 'onfig' 
    $caentries = $cadump -split [System.Environment]::NewLine.ToCharArray()
    if ($caentries[0]) {
    $caserver1 = [regex]::match($caentries[0], '"([^"]+)"').Groups[1].Value
    } else {}
    if ($caentries[1]) {
    $caserver2 = [regex]::match($caentries[1], '"([^"]+)"').Groups[1].Value
    } else {}
    if ($caentries[2]) {
    $caserver3 = [regex]::match($caentries[2], '"([^"]+)"').Groups[1].Value
    } else {}
    if ($caentries[3]) {
    $caserver4 = [regex]::match($caentries[3], '"([^"]+)"').Groups[1].Value
    } else {}
    if ($caentries[4]) {
    $caserver5 = [regex]::match($caentries[4], '"([^"]+)"').Groups[1].Value
    } else {}

    if ($cadump.Count -eq 1) {$calist = $caserver1}
    elseif ($cadump.Count -eq 2) {$calist = $caserver1,$caserver2}
    elseif ($cadump.Count -eq 3) {$calist = $caserver1,$caserver2,$caserver3}
    elseif ($cadump.Count -eq 4) {$calist = $caserver1,$caserver2,$caserver3,$caserver4}
    elseif ($cadump.Count -eq 5) {$calist = $caserver1,$caserver2,$caserver3,$caserver4,$caserver5}
    else {}

# implementing Get-Folder function for output-directory picker #

Function Get-Folder($initialDirectory) {
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $FolderBrowserDialog.RootFolder = 'MyComputer'
        if ($initialDirectory) { $FolderBrowserDialog.SelectedPath = $initialDirectory }
        [void] $FolderBrowserDialog.ShowDialog()
        return $FolderBrowserDialog.SelectedPath
        }

# setting template entries for gui #
# These are examples - put in the correct certificate template names from your ca server #
$templates = "Webserver","Printer","Computer","Laptop","Thinclient","MobileDevice","CodeSignature","User"

# GUI script part begins here #

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$objForm = New-Object System.Windows.Forms.Form
#$objForm.Backcolor=""
$objForm.StartPosition = "CenterScreen"
$objForm.Size = New-Object System.Drawing.Size(730,540)
$objForm.Text = "Create CSR / Export Certificate"
#---infotext/textbox---#
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,10)
$objLabel.Size = New-Object System.Drawing.Size(400,20)
$objLabel.Text = "Enter Hostname (full qualified) here:"
$objForm.Controls.Add($objLabel)
#---text-box---#
$objHostname = New-Object System.Windows.Forms.TextBox
$objHostname.Location = New-Object System.Drawing.Size(11,30)
$objHostname.Size = New-Object System.Drawing.Size(400,60)
$objForm.Controls.Add($objHostname)
#---infotext/templates---#
$objTemplateLabel = New-Object System.Windows.Forms.Label
$objTemplateLabel.Location = New-Object System.Drawing.Size(10,70)
$objTemplateLabel.Size = New-Object System.Drawing.Size(400,20)
$objTemplateLabel.Text = "Select Certificate Template:"
$objForm.Controls.Add($objTemplateLabel)
#---templates-dropdown-menu---#
$objTemplates = New-Object System.Windows.Forms.Combobox
$objTemplates.Location = New-Object System.Drawing.Size(11,90)
$objTemplates.Size = New-Object System.Drawing.Size(400,60)
$objTemplates.Height = 70
$objForm.Controls.Add($objTemplates)
$objForm.Topmost = $True
$objForm.Add_Shown({$objForm.Activate()})
$objTemplates.Items.AddRange($templates) #templates loaded from variable
$objTemplates.SelectedItem #chosen template will be applied
#---infotext/SAN-entries---#
$objSANLabel = New-Object System.Windows.Forms.Label
$objSANLabel.Location = New-Object System.Drawing.Size(10,130)
$objSANLabel.Size = New-Object System.Drawing.Size(400,40)
$objSANLabel.Text = "Enter SAN's (subject alternative names) here.`nSeparating character is comma. (,) `nOnly DNS-entries and aliases allowed - maximum entries: 20"
$objForm.Controls.Add($objSANLabel)
#---SAN-entries---#
$objSANs = New-Object System.Windows.Forms.TextBox
$objSANs.Location = New-Object System.Drawing.Size(11,180)
$objSANs.Size = New-Object System.Drawing.Size(400,60)
$objForm.Controls.Add($objSANs)
#---infotext/issuingCA---#
$objCALabel = New-Object System.Windows.Forms.Label
$objCALabel.Location = New-Object System.Drawing.Size(10,210)
$objCALabel.Size = New-Object System.Drawing.Size(400,20)
$objCALabel.Text = "Select issuing CA:"
$objForm.Controls.Add($objCALabel)
#---issuingCA---#
$objCAs = New-Object System.Windows.Forms.Combobox
$objCAs.Location = New-Object System.Drawing.Size(11,230)
$objCAs.Size = New-Object System.Drawing.Size(400,60)
$objCAs.Height = 70
$objForm.Controls.Add($objCAs)
$objForm.Topmost = $True
$objForm.Add_Shown({$objForm.Activate()})
$objCAs.Items.AddRange($calist) #Issuing CAs will be loaded from variable and listed
$objCAs.SelectedItem #select CA will now be applied
#---infotext/own passphrase---#
$objownpwdLabel = New-Object System.Windows.Forms.Label
$objownpwdLabel.Location = New-Object System.Drawing.Size(440,320)
$objownpwdLabel.Size = New-Object System.Drawing.Size(240,60)
$objownpwdLabel.Text = "Enter a passphrase.`nIf empty, the script will use the hostname (without domain suffix) spelled backwards."
$objForm.Controls.Add($objownpwdLabel)
#---own passphrase---#
$objownpwd = New-Object System.Windows.Forms.TextBox
$objownpwd.Location = New-Object System.Drawing.Size(440,380)
$objownpwd.Size = New-Object System.Drawing.Size(240,80)
$objForm.Controls.Add($objownpwd)
#--Infotext-output-folder--#
$objOutDir = New-Object System.Windows.Forms.Label
$objOutDir.Location = New-Object System.Drawing.Size(10,400)
$objOutDir.Size = New-Object System.Drawing.Size(400,20)
$objOutDir.Text = "Select output directory:"
$objForm.Controls.Add($objOutDir)
#--Output-folder--#
$BrowseButton = New-Object System.Windows.Forms.Button
$BrowseButton.Location = New-Object System.Drawing.Size(10,420)
$BrowseButton.Size = New-Object System.Drawing.Size(75,23)
$BrowseButton.Text = "Browse..."
$BrowseButton.Name = "Browse"

$objOutDirLabel = New-Object System.Windows.Forms.Label
$objOutDirLabel.Location = New-Object System.Drawing.Size(10,470)
$objOutDirLabel.Size = New-Object System.Drawing.Size(400,20) 
$objForm.Controls.Add($objOutDirLabel) 

$objDirLabel = New-Object System.Windows.Forms.Label
$objDirLabel.Location = New-Object System.Drawing.Size(10,450)
$objDirLabel.Size = New-Object System.Drawing.Size(400,20)
#$objDirLabel.Text = "-click `"Browse`" to select-"
$objForm.Controls.Add($objDirLabel)
#--browse-button--#
$BrowseButton.Add_Click({
    $global:outdir = Get-Folder
    $objDirLabel.Text = $outdir
    })

$objForm.Controls.Add($BrowseButton) 


#---Infotext/OpenSSL---#
$objOpenSSL = New-Object System.Windows.Forms.Label
$objOpenSSL.Location = New-Object System.Drawing.Size(10,300)
$objOpenSSL.Size = New-Object System.Drawing.Size(400,70)
$objOpenSSL.Text = "Note: 
1.OpenSSL for Windows must be installed in setup-default directory.
2.Make sure, you have sufficient rights on your CA-server
3.Also make sure, you have write access in your destination directory.
If the operation fails, check Windows-EventLog on your CA-server"
$objForm.Controls.Add($objOpenSSL)

#---Infotext/Logo---#
# this is just a flex. change the text to whatever you want :) #
$objLogo = New-Object System.Windows.Forms.Label
$objLogo.Location = New-Object System.Drawing.Size(530,60)
$objLogo.Size = New-Object System.Drawing.Size(600,200)
$objLogo.Text = " \|/
-O-__________
 /|\'.....................\
 |'.......New.........|
 |........Cert.........|
 |........................|
 |........................|
 |........................|
 |...........by  P.R.|
 |____________|
"
$objForm.Controls.Add($objLogo)

#--Create/Export button--#
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(440,440)
$OKButton.Size = New-Object System.Drawing.Size(90,23)
$OKButton.Text = "Create/Export"
$OKButton.Name = "Create/Export"
$OKButton.Add_Click({

    #--- CSR data and SANs will be applied here---#

    $global:name = $objHostname.Text
    $global:shortname = $name.Substring(0, $name.IndexOf('.'))

    $global:sans = $objSANs.Text.Split(",")
    $global:san1 = $sans[0]
    $global:san2 = $sans[1]
    $global:san3 = $sans[2]
    $global:san4 = $sans[3]
    $global:san5 = $sans[4]
    $global:san6 = $sans[5]
    $global:san7 = $sans[6]
    $global:san8 = $sans[7]
    $global:san9 = $sans[8]
    $global:san10 = $sans[9]
    $global:san11 = $sans[10]
    $global:san12 = $sans[11]
    $global:san13 = $sans[12]
    $global:san14 = $sans[13]
    $global:san15 = $sans[14]
    $global:san16 = $sans[15]
    $global:san17 = $sans[16]
    $global:san18 = $sans[17]
    $global:san19 = $sans[18]
    $global:san20 = $sans[19]

    # The script is creating now a CSR (Certificate Signing Request)
    $csrPath = Join-Path -Path "$outdir" -ChildPath "$shortname.csr"
    $infPath = Join-Path -Path "$outdir" -ChildPath "$shortname.inf"
    $cerPath = $csrPath -replace ".csr" -replace ""

    # Here begins the INF-file content part, which is necessary for the CSR.
    # If necessary, you can change the values like key length etc 

    $global:infContents =
@"
[Version]
Signature = "`$Windows NT`$"

[NewRequest]
Subject = "CN=$name"
KeySpec = 1
KeyLength = 4096
Exportable = TRUE
MachineKeySet = TRUE
SMIME = False
PrivateKeyArchive = FALSE
UserProtected = FALSE
UseExistingKeySet = FALSE
ProviderName = "Microsoft RSA SChannel Cryptographic Provider"
HashAlgorithm = SHA256
ProviderType = 12
RequestType = PKCS10
KeyUsage = 0xa0

[EnhancedKeyUsageExtension]
OID=1.3.6.1.5.5.7.3.1

[Extensions]
2.5.29.17 = "{text}"
_continue_ = "dns=$san1&"
_continue_ = "dns=$san2&"
_continue_ = "dns=$san3&"
_continue_ = "dns=$san4&"
_continue_ = "dns=$san5&"
_continue_ = "dns=$san6&"
_continue_ = "dns=$san7&"
_continue_ = "dns=$san8&"
_continue_ = "dns=$san9&"
_continue_ = "dns=$san10&"
_continue_ = "dns=$san11&"
_continue_ = "dns=$san12&"
_continue_ = "dns=$san13&"
_continue_ = "dns=$san14&"
_continue_ = "dns=$san15&"
_continue_ = "dns=$san16&"
_continue_ = "dns=$san17&"
_continue_ = "dns=$san18&"
_continue_ = "dns=$san19&"
_continue_ = "dns=$san20"
"@

        
$global:ca = $objCAs.Text
$global:caserver = $ca.Substring(0, $ca.IndexOf('\'))
$global:template = $objTemplates.Text

# INF-file will now be saved and CSR will be created.
$infContents | Out-File -filepath $infPath -force
Invoke-Command -ComputerName $caserver -ScriptBlock {
    if (!(Test-Path C:\Temp))
    {
    New-Item -itemType Directory -Path C:\Temp
    }else
    {}
    }
Get-Item -Path "$infPath" | Move-Item -Destination "\\$caserver\C$\Temp"

Invoke-Command -ComputerName $caserver -ScriptBlock {
    certreq -new "C:\Temp\$Using:shortname.inf" "C:\Temp\$Using:shortname.csr"
    }
# creating target directory
$targetdir = "$outdir"+"\"+"$shortname"
New-Item $targetdir -ItemType Directory

# the CSR will be signed by your CA, the received certificate will be saved in the destination directory
        
Invoke-Command -ComputerName $caserver -ScriptBlock { 
    certreq -submit -config "$using:ca" -attrib certificateTemplate:$Using:template "C:\Temp\$using:shortname.csr" "C:\Temp\$using:shortname.cer" 
    }

# Importing the just created certificate to the ca-server under "my certificates".
# You don't have to, but this is helpful if you monitor the certificates on your CA server to prevent them from expiring.
# After then, key and cert will be exported as PFX to the destination folder.
# Using the thumbprint (is unique), the certificate can be retrieved from the server.
        
$global:thumbprint = Invoke-Command -ComputerName $caserver -ScriptBlock { Import-Certificate -FilePath "C:\Temp\$using:shortname.cer" -Confirm:$false -CertStoreLocation "Cert:\LocalMachine\My" | Select-Object Thumbprint }
$global:thumb = $thumbprint.Thumbprint

# Passphrase = hostname written backwards. 
# The passphrase will now be converted to a secure-string.
# For OpenSSL, the clear-text password will be used, because OpenSSL doesn't recognize a secure-string.
# If user typed in an own passphrase, it will be used.

if (!$objownpwd.Text) {
    $global:pwd = $shortname
    $plainpwd = $pwd[-1..-$pwd.Length] -join ''
    $global:securepwd = convertto-securestring -String $plainpwd -asplaintext -force
    }
    else
    {
    $global:ownpwd = $objownpwd.Text
    $global:plainpwd = $ownpwd
    $global:securepwd = convertto-securestring -String $plainpwd -asplaintext -force
    $global:pwd = $ownpwd
    }

# The PFX file is now being saved and protected with the passphrase.
$global:hostname = $env:computername
$global:remdir = $outdir.Replace(':','$')

               
Invoke-Command -ComputerName $caserver -ScriptBlock {
    Get-ChildItem -Path "Cert:\LocalMachine\my\$using:thumb" | Export-PfxCertificate -FilePath "C:\Temp\$using:shortname.pfx" -Password $using:securepwd
    }
Get-Item -Path "\\$caserver\C$\Temp\$shortname.pfx" | Move-Item -Destination "$outdir\$shortname\$shortname.pfx"
Get-Item -Path "\\$caserver\C$\Temp\$shortname.cer" | Move-Item -Destination "$outdir\$shortname\$shortname.cer"

if (Test-Path "C:\Program Files\OpenSSL-Win64\bin\openssl.exe")
    {
    $openssldir = "C:\Program Files\OpenSSL-Win64\bin\openssl.exe"
    } 
    else {
    $openssldir = "C:\Program Files (x86)\OpenSSL-Win32\bin\openssl.exe"
    }

        

# OpenSSL splits the certificate and private key so we have two seperate files: The key (.KEY-file) and the certificate (.CER-Datei).
cmd /c "C:\Program Files\OpenSSL-Win64\bin\openssl.exe" pkcs12 -in "$outdir\$shortname\$shortname.pfx"  -out "$outdir\$shortname\$shortname.tmp"  -passin pass:$plainpwd -passout pass:$plainpwd
cmd /c "C:\Program Files\OpenSSL-Win64\bin\openssl.exe" pkey -in "$outdir\$shortname\$shortname.tmp" -out "$outdir\$shortname\$shortname.key" -passin pass:$plainpwd
Remove-Item -Force "$outdir\$shortname\$shortname.tmp"

[void] [Windows.Forms.MessageBox]::Show("Certificate created, export Completed. `nCheck Output Directory.")
})
$objForm.Controls.Add($OKButton) 

#--Cancel-Button--#
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(605,440)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Quit"
$CancelButton.Name = "Quit"
$CancelButton.DialogResult = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton) 



[void] $objForm.ShowDialog()
