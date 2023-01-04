# ---------------------------------------------------------------------------------------------#
# This script creates a certificate with subject saternative names (DNS), using a CSR          #
# imports it into the Windows certificate storage (my certificates) on your issuing CA-server  #
# exports the public certificate as CER-File,                                                  #
# the public certificate including private key as PFX-File and                                 #
# the private key as KEY-File to your destination directory .                                  #
# ---------------------------------------------------------------------------------------------#

# checking if OpenSSL is installed #
if (!(Test-Path "C:\Program Files\OpenSSL-Win64\bin\openssl.exe") -or (Test-Path "C:\Program Files (x86)\OpenSSL-Win32\bin\openssl.exe"))
    {
    [System.Windows.MessageBox]::Show('OpenSSL for Windows was not found in the setup-default directory. Please install OpenSSL first!','Warning')
    } 
    else {}


# checking current locale #

$locale = Get-WinSystemLocale
if ($locale.Name -eq "de-AT")
            {
            $cadump = certutil -dump | select-string -Pattern Konfiguration 
            $caentries = $cadump -replace "  Konfiguration:          	" -replace ""
            }
            else
            {}

if ($locale.Name -eq "de-DE")
            {
            $cadump = certutil -dump | select-string -Pattern Konfiguration 
            $caentries = $cadump -replace "  Konfiguration:          	" -replace ""
            }
            else
            {}
if ($locale.Name -eq "en-US")
            {
            $cadump = certutil -dump | select-string -Pattern configuration 
            $caentries = $cadump -replace "  configuration:          	" -replace ""
            }
            else
            {}

$calist = $caentries -replace "`"" -replace ""

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
$templates = "AIX","WebserverPVA_1J","PrinterPVA","ComputerPVA","DevicePVA","LaptopPVA","ProtelPVA","ThinclientPVA","MobileDevicePVA","CodesignaturPVA","Benutzer"

# GUI script part begins here #

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$objForm = New-Object System.Windows.Forms.Form
#$objForm.Backcolor=""
$objForm.StartPosition = "CenterScreen"
$objForm.Size = New-Object System.Drawing.Size(730,560)
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
#--browse-button--#
$BrowseButton.Add_Click({
    $outdir = Get-Folder
    $objOutDirLabel.Text = $outdir
    })

$objForm.Controls.Add($BrowseButton) 
$objDirLabel = New-Object System.Windows.Forms.Label
$objDirLabel.Location = New-Object System.Drawing.Size(10,450)
$objDirLabel.Size = New-Object System.Drawing.Size(400,20)
$objDirLabel.Text = "Selected output directory:"
$objForm.Controls.Add($objDirLabel)

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

#---Infotext/creator---#
$objOpenSSL = New-Object System.Windows.Forms.Label
$objOpenSSL.Location = New-Object System.Drawing.Size(530,60)
$objOpenSSL.Size = New-Object System.Drawing.Size(600,200)
$objOpenSSL.Text = " \|/
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
$objForm.Controls.Add($objOpenSSL)

#--Create/Export Button--#
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(500,470)
$OKButton.Size = New-Object System.Drawing.Size(85,23)
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

    # The script is creating now an CSR (Certificate Signing Request)
    $csrPath = Join-Path -Path "$outdir" -ChildPath "$shortname.csr"
    $infPath = Join-Path -Path "$outdir" -ChildPath "$shortname.inf"
    $cerPath = $csrPath -replace ".csr" -replace ""

    # Here begins the INF-file content art, which is necessary for the CSR.

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
New-Item "$outdir\$shortname" -ItemType Directory

# certificate request (CSR) will be signed by your CA, the received certificate will be saved in the destination directory
        
Invoke-Command -ComputerName $caserver -ScriptBlock { 
    certreq -submit -config "$using:ca" -attrib certificateTemplate:$Using:template "C:\Temp\$using:shortname.csr" "C:\Temp\$using:shortname.cer" 
    }

# Importing the just created certificate to the ca-server under "my certificates".
# After then, key and cert will be exported as PFX to the destination folder.
# Using the thumbprint (is unique), ther certificate can be retrieved from the server.
        
$global:thumbprint = Invoke-Command -ComputerName $caserver -ScriptBlock { Import-Certificate -FilePath "C:\Temp\$using:shortname.cer" -Confirm:$false -CertStoreLocation "Cert:\LocalMachine\My" | Select-Object Thumbprint }
$global:thumb = $thumbprint.Thumbprint

# Passphrase = hostname written backwards. 
# The Passphrase will now be converted to a secure-string.
# For OpenSSL, the clear-text password will be used, because OpenSSL doesn't recognize a secure-string.
$global:pwd = $shortname
$plainpwd = $pwd[-1..-$pwd.Length] -join ''
$global:securepwd = convertto-securestring -String $plainpwd -asplaintext -force

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
$CancelButton.Location = New-Object System.Drawing.Size(600,470)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Quit"
$CancelButton.Name = "Quit"
$CancelButton.DialogResult = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton) 



[void] $objForm.ShowDialog()




# 2022 / Rezanka / HREZ
# SIG # Begin signature block
# MIIl3gYJKoZIhvcNAQcCoIIlzzCCJcsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUx7/HdYm4BU+bOHhsf+rdMxHR
# IQmggh/1MIIFjTCCBHWgAwIBAgIQDpsYjvnQLefv21DiCEAYWjANBgkqhkiG9w0B
# AQwFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMjIwODAxMDAwMDAwWhcNMzExMTA5MjM1OTU5WjBiMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQw
# ggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC/5pBzaN675F1KPDAiMGkz
# 7MKnJS7JIT3yithZwuEppz1Yq3aaza57G4QNxDAf8xukOBbrVsaXbR2rsnnyyhHS
# 5F/WBTxSD1Ifxp4VpX6+n6lXFllVcq9ok3DCsrp1mWpzMpTREEQQLt+C8weE5nQ7
# bXHiLQwb7iDVySAdYyktzuxeTsiT+CFhmzTrBcZe7FsavOvJz82sNEBfsXpm7nfI
# SKhmV1efVFiODCu3T6cw2Vbuyntd463JT17lNecxy9qTXtyOj4DatpGYQJB5w3jH
# trHEtWoYOAMQjdjUN6QuBX2I9YI+EJFwq1WCQTLX2wRzKm6RAXwhTNS8rhsDdV14
# Ztk6MUSaM0C/CNdaSaTC5qmgZ92kJ7yhTzm1EVgX9yRcRo9k98FpiHaYdj1ZXUJ2
# h4mXaXpI8OCiEhtmmnTK3kse5w5jrubU75KSOp493ADkRSWJtppEGSt+wJS00mFt
# 6zPZxd9LBADMfRyVw4/3IbKyEbe7f/LVjHAsQWCqsWMYRJUadmJ+9oCw++hkpjPR
# iQfhvbfmQ6QYuKZ3AeEPlAwhHbJUKSWJbOUOUlFHdL4mrLZBdd56rF+NP8m800ER
# ElvlEFDrMcXKchYiCd98THU/Y+whX8QgUWtvsauGi0/C1kVfnSD8oR7FwI+isX4K
# Jpn15GkvmB0t9dmpsh3lGwIDAQABo4IBOjCCATYwDwYDVR0TAQH/BAUwAwEB/zAd
# BgNVHQ4EFgQU7NfjgtJxXWRM3y5nP+e6mK4cD08wHwYDVR0jBBgwFoAUReuir/SS
# y4IxLVGLp6chnfNtyA8wDgYDVR0PAQH/BAQDAgGGMHkGCCsGAQUFBwEBBG0wazAk
# BggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAC
# hjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURS
# b290Q0EuY3J0MEUGA1UdHwQ+MDwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwEQYDVR0gBAowCDAGBgRV
# HSAAMA0GCSqGSIb3DQEBDAUAA4IBAQBwoL9DXFXnOF+go3QbPbYW1/e/Vwe9mqyh
# hyzshV6pGrsi+IcaaVQi7aSId229GhT0E0p6Ly23OO/0/4C5+KH38nLeJLxSA8hO
# 0Cre+i1Wz/n096wwepqLsl7Uz9FDRJtDIeuWcqFItJnLnU+nBgMTdydE1Od/6Fmo
# 8L8vC6bp8jQ87PcDx4eo0kxAGTVGamlUsLihVo7spNU96LHc/RzY9HdaXFSMb++h
# UD38dglohJ9vytsgjTVgHAIDyyCwrFigDkBjxZgiwbJZ9VVrzyerbHbObyMt9H5x
# aiNrIv8SuFQtJ37YOtnwtoeW/VvRXKwYw02fc7cBqZ9Xql4o4rmUMIIFpDCCA4yg
# AwIBAgITUwAAAASev5oh1tKVZQAAAAAABDANBgkqhkiG9w0BAQsFADAuMSwwKgYD
# VQQDEyNQZW5zaW9uc3ZlcnNpY2hlcnVuZ3NhbnN0YWx0IFJvb3RDQTAeFw0yMDAy
# MjYxMzUxNThaFw0yODAyMjYxNDAxNThaMHYxEjAQBgoJkiaJk/IsZAEZFgJhdDEX
# MBUGCgmSJomT8ixkARkWB3NvenZlcnMxEzARBgoJkiaJk/IsZAEZFgNwdmExMjAw
# BgNVBAMTKVBlbnNpb25zdmVyc2ljaGVydW5nc2Fuc3RhbHQgU3ViQ0FEZXZpY2Vz
# MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAp3pX/R+3NWnAlxQrVUxA
# Nuuxd3qoDYlDAkYvT9nxySxFfx8ybv+ELVBsgfHp3xUQKMAXKLGNSqm4qkoVxnAm
# 1mVguMUxpTHgJfS5uacsFU/kpox0y6rSfyRcTqI9bxvFwhpUaUETN5djooU8rvcE
# 4JMEAevVibtk4xxPAYlkuqKjb2aCvxuTAnENxP95VCH4JIzxxave/wmL5BbHAKdj
# WC6w6aULUvZU3Bkum6/c+ZRxxnF589l8c95x/GuYQMJjrbiG+jNs03IVf/h/x36p
# bAF2RLJwx5AiyzymDPLkAx3GzKihXSZLzsskoEb8kFHj0N7pt5mY1BQbkqbymrov
# RQIDAQABo4IBcTCCAW0wEAYJKwYBBAGCNxUBBAMCAQEwIwYJKwYBBAGCNxUCBBYE
# FMpE51uOvWwLWQqEhDBy4Xt51zjKMB0GA1UdDgQWBBRCVcrcsCyGX2gORX1g5y05
# RC2mkDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBQHb4kNHbzfuGdF62egX2ULpIM23jBX
# BgNVHR8EUDBOMEygSqBIhkZodHRwOi8vcm9vdGNhLnB2YS5zb3p2ZXJzLmF0L1Bl
# bnNpb25zdmVyc2ljaGVydW5nc2Fuc3RhbHQlMjBSb290Q0EuY3JsMGIGCCsGAQUF
# BwEBBFYwVDBSBggrBgEFBQcwAoZGaHR0cDovL3Jvb3RjYS5wdmEuc296dmVycy5h
# dC9QZW5zaW9uc3ZlcnNpY2hlcnVuZ3NhbnN0YWx0JTIwUm9vdENBLmNydDANBgkq
# hkiG9w0BAQsFAAOCAgEAZkOVMhnK/qSNDtZBT6ZoWlCtY26f3pRMxXfIwKdsCh59
# 40v9/TB7T1ZRcebwEipDKuV5ywGa2j1PyfSh0EyXNHsxVeEZmsiocv96g5xBXUsy
# Ha8WVqbWdNBgVv24wuYK+YUyd517S8PKQgx4Qjmh1B4DMGq9WTvjaNxi3H5fhfHs
# LxsLb5bwr/y07qrNNNWvRP96d13dtk6w/tJSvYOWsqw+St64EotiNiJRWV4dvG/M
# VQ4I/VlT3QHavancWWnyYNO7IjCupwt2IiOhHryoBSsMxAEB+ak1Ga2JXL03Hd/D
# hFctpIWrNpXFft5rQHwezRT+tkP++EaGEecM9ViQD/yrjz1u77TGRQljwUAtSc3L
# JQOuZFV1NyZAvIHp5uOeZcMMMyWX2ApccEDqSRrC9CUWhBnvX6G0k/oNJV917qVI
# djYVEFcZvoegsxC50k2nPJWCvvdNE+ErpF2kiz+YmuuRLJxd2wPcNBCG5RfPPyVE
# Z40uJ0AYKIDmrk6XaGWf2a0pEtS6JmQaIXvUxb4ugIcIFOho4RB1euEdAV1VdwqZ
# Jz0iV5+YX2HVjpRTfBa+KcAYV6hccKQjD2Seq5odASJqvMpm4RZUc+Pg85wbjxGg
# CIg2FmTzbigLiIE75rePlpS+6wf7B/19hS3P5IfDPLlt+yqSidfz71mEcj2ip44w
# ggauMIIElqADAgECAhAHNje3JFR82Ees/ShmKl5bMA0GCSqGSIb3DQEBCwUAMGIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBH
# NDAeFw0yMjAzMjMwMDAwMDBaFw0zNzAzMjIyMzU5NTlaMGMxCzAJBgNVBAYTAlVT
# MRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1
# c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0EwggIiMA0GCSqG
# SIb3DQEBAQUAA4ICDwAwggIKAoICAQDGhjUGSbPBPXJJUVXHJQPE8pE3qZdRodbS
# g9GeTKJtoLDMg/la9hGhRBVCX6SI82j6ffOciQt/nR+eDzMfUBMLJnOWbfhXqAJ9
# /UO0hNoR8XOxs+4rgISKIhjf69o9xBd/qxkrPkLcZ47qUT3w1lbU5ygt69OxtXXn
# HwZljZQp09nsad/ZkIdGAHvbREGJ3HxqV3rwN3mfXazL6IRktFLydkf3YYMZ3V+0
# VAshaG43IbtArF+y3kp9zvU5EmfvDqVjbOSmxR3NNg1c1eYbqMFkdECnwHLFuk4f
# sbVYTXn+149zk6wsOeKlSNbwsDETqVcplicu9Yemj052FVUmcJgmf6AaRyBD40Nj
# gHt1biclkJg6OBGz9vae5jtb7IHeIhTZgirHkr+g3uM+onP65x9abJTyUpURK1h0
# QCirc0PO30qhHGs4xSnzyqqWc0Jon7ZGs506o9UD4L/wojzKQtwYSH8UNM/STKvv
# mz3+DrhkKvp1KCRB7UK/BZxmSVJQ9FHzNklNiyDSLFc1eSuo80VgvCONWPfcYd6T
# /jnA+bIwpUzX6ZhKWD7TA4j+s4/TXkt2ElGTyYwMO1uKIqjBJgj5FBASA31fI7tk
# 42PgpuE+9sJ0sj8eCXbsq11GdeJgo1gJASgADoRU7s7pXcheMBK9Rp6103a50g5r
# mQzSM7TNsQIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4E
# FgQUuhbZbU2FL3MpdpovdYxqII+eyG8wHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5n
# P+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMIMHcG
# CCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQu
# Y29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGln
# aUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8v
# Y3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAgBgNV
# HSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIB
# AH1ZjsCTtm+YqUQiAX5m1tghQuGwGC4QTRPPMFPOvxj7x1Bd4ksp+3CKDaopafxp
# wc8dB+k+YMjYC+VcW9dth/qEICU0MWfNthKWb8RQTGIdDAiCqBa9qVbPFXONASIl
# zpVpP0d3+3J0FNf/q0+KLHqrhc1DX+1gtqpPkWaeLJ7giqzl/Yy8ZCaHbJK9nXzQ
# cAp876i8dU+6WvepELJd6f8oVInw1YpxdmXazPByoyP6wCeCRK6ZJxurJB4mwbfe
# Kuv2nrF5mYGjVoarCkXJ38SNoOeY+/umnXKvxMfBwWpx2cYTgAnEtp/Nh4cku0+j
# Sbl3ZpHxcpzpSwJSpzd+k1OsOx0ISQ+UzTl63f8lY5knLD0/a6fxZsNBzU+2QJsh
# IUDQtxMkzdwdeDrknq3lNHGS1yZr5Dhzq6YBT70/O3itTK37xJV77QpfMzmHQXh6
# OOmc4d0j/R0o08f56PGYX/sr2H7yRp11LB4nLCbbbxV7HhmLNriT1ObyF5lZynDw
# N7+YAN8gFk8n+2BnFqFmut1VwDophrCYoCvtlUG3OtUVmDG0YgkPCr2B2RP+v6TR
# 81fZvAT6gt4y3wSJ8ADNXcL50CN/AAvkdgIm2fBldkKmKYcJRyvmfxqkhQ/8mJb2
# VVQrH4D6wPIOK+XW+6kvRBVK5xMOHds3OBqhK/bt1nz8MIIGwDCCBKigAwIBAgIQ
# DE1pckuU+jwqSj0pB4A9WjANBgkqhkiG9w0BAQsFADBjMQswCQYDVQQGEwJVUzEX
# MBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0
# ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBMB4XDTIyMDkyMTAw
# MDAwMFoXDTMzMTEyMTIzNTk1OVowRjELMAkGA1UEBhMCVVMxETAPBgNVBAoTCERp
# Z2lDZXJ0MSQwIgYDVQQDExtEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMiAtIDIwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDP7KUmOsap8mu7jcENmtuh6BSF
# dDMaJqzQHFUeHjZtvJJVDGH0nQl3PRWWCC9rZKT9BoMW15GSOBwxApb7crGXOlWv
# M+xhiummKNuQY1y9iVPgOi2Mh0KuJqTku3h4uXoW4VbGwLpkU7sqFudQSLuIaQyI
# xvG+4C99O7HKU41Agx7ny3JJKB5MgB6FVueF7fJhvKo6B332q27lZt3iXPUv7Y3U
# TZWEaOOAy2p50dIQkUYp6z4m8rSMzUy5Zsi7qlA4DeWMlF0ZWr/1e0BubxaompyV
# R4aFeT4MXmaMGgokvpyq0py2909ueMQoP6McD1AGN7oI2TWmtR7aeFgdOej4TJEQ
# ln5N4d3CraV++C0bH+wrRhijGfY59/XBT3EuiQMRoku7mL/6T+R7Nu8GRORV/zbq
# 5Xwx5/PCUsTmFntafqUlc9vAapkhLWPlWfVNL5AfJ7fSqxTlOGaHUQhr+1NDOdBk
# +lbP4PQK5hRtZHi7mP2Uw3Mh8y/CLiDXgazT8QfU4b3ZXUtuMZQpi+ZBpGWUwFjl
# 5S4pkKa3YWT62SBsGFFguqaBDwklU/G/O+mrBw5qBzliGcnWhX8T2Y15z2LF7OF7
# ucxnEweawXjtxojIsG4yeccLWYONxu71LHx7jstkifGxxLjnU15fVdJ9GSlZA076
# XepFcxyEftfO4tQ6dwIDAQABo4IBizCCAYcwDgYDVR0PAQH/BAQDAgeAMAwGA1Ud
# EwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwIAYDVR0gBBkwFzAIBgZn
# gQwBBAIwCwYJYIZIAYb9bAcBMB8GA1UdIwQYMBaAFLoW2W1NhS9zKXaaL3WMaiCP
# nshvMB0GA1UdDgQWBBRiit7QYfyPMRTtlwvNPSqUFN9SnDBaBgNVHR8EUzBRME+g
# TaBLhklodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRS
# U0E0MDk2U0hBMjU2VGltZVN0YW1waW5nQ0EuY3JsMIGQBggrBgEFBQcBAQSBgzCB
# gDAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMFgGCCsGAQUF
# BzAChkxodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVk
# RzRSU0E0MDk2U0hBMjU2VGltZVN0YW1waW5nQ0EuY3J0MA0GCSqGSIb3DQEBCwUA
# A4ICAQBVqioa80bzeFc3MPx140/WhSPx/PmVOZsl5vdyipjDd9Rk/BX7NsJJUSx4
# iGNVCUY5APxp1MqbKfujP8DJAJsTHbCYidx48s18hc1Tna9i4mFmoxQqRYdKmEIr
# UPwbtZ4IMAn65C3XCYl5+QnmiM59G7hqopvBU2AJ6KO4ndetHxy47JhB8PYOgPvk
# /9+dEKfrALpfSo8aOlK06r8JSRU1NlmaD1TSsht/fl4JrXZUinRtytIFZyt26/+Y
# siaVOBmIRBTlClmia+ciPkQh0j8cwJvtfEiy2JIMkU88ZpSvXQJT657inuTTH4YB
# ZJwAwuladHUNPeF5iL8cAZfJGSOA1zZaX5YWsWMMxkZAO85dNdRZPkOaGK7DycvD
# +5sTX2q1x+DzBcNZ3ydiK95ByVO5/zQQZ/YmMph7/lxClIGUgp2sCovGSxVK05iQ
# RWAzgOAj3vgDpPZFR+XOuANCR+hBNnF3rf2i6Jd0Ti7aHh2MWsgemtXC8MYiqE+b
# vdgcmlHEL5r2X6cnl7qWLoVXwGDneFZ/au/ClZpLEQLIgpzJGgV8unG1TnqZbPTo
# ntRamMifv427GFxD9dAq6OJi7ngE273R+1sKqHB+8JeEeOMIA11HLGOoJTiXAdI/
# Otrl5fbmm9x+LMz/F0xNAKLY1gEOuIvu5uByVYksJxlh9ncBjDCCB0IwggYqoAMC
# AQICE0oAAI6eyuHJiZ0aHnkAAQAAjp4wDQYJKoZIhvcNAQELBQAwdjESMBAGCgmS
# JomT8ixkARkWAmF0MRcwFQYKCZImiZPyLGQBGRYHc296dmVyczETMBEGCgmSJomT
# 8ixkARkWA3B2YTEyMDAGA1UEAxMpUGVuc2lvbnN2ZXJzaWNoZXJ1bmdzYW5zdGFs
# dCBTdWJDQURldmljZXMwHhcNMjAxMTI3MTIyMTMxWhcNMjMxMTI3MTIyMTMxWjCB
# gjELMAkGA1UEBhMCQVQxJTAjBgNVBAoTHFBlbnNpb25zdmVyc2ljaGVydW5nc2Fu
# c3RhbHQxJTAjBgNVBAsTHFBlbnNpb25zdmVyc2ljaGVydW5nc2Fuc3RhbHQxJTAj
# BgNVBAMTHFBlbnNpb25zdmVyc2ljaGVydW5nc2Fuc3RhbHQwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQC34XFtuPbhDR5sf/J2EjJ+WbHIZzCTKTecCjDQ
# FP7yP4cuxDIzxH+9CJA9qT9EqOFVAOMdYki4FXGo7VwSEXz38V3v45CNQ1qmZ48J
# NC/D1nBFa0kUYTSXiyUq0S+K7hbdAvRCcqWEG+0L4NrIQbFNrG95TGVrrErsVrBk
# +zp3fOJxilwZCnpuzRMErrHLK8X7Ls/qcnjwiuMDL+69a5/0yslqxJBRXRN2otwS
# 2NBs9UXXXLj64+FSdk4N+RzBsJ23UBxO2oRVqBRHcgZwqp1of9wTEnWhzWkpUdPo
# uyt84zzbpCaW11a0vsiXRwrzj2nBIheODZgG3nVM65uDU9g1AgMBAAGjggO6MIID
# tjA9BgkrBgEEAYI3FQcEMDAuBiYrBgEEAYI3FQiCpfJIg/6BU4OtizWBiMhogdTp
# MwqFnPIUhdzTSAIBZQIBBDAOBgNVHQ8BAf8EBAMCBLAwOwYDVR0lBDQwMgYIKwYB
# BQUHAwgGCCsGAQUFBwMEBggrBgEFBQcDAQYIKwYBBQUHAwMGCCsGAQUFBwMCMEsG
# CSsGAQQBgjcVCgQ+MDwwCgYIKwYBBQUHAwgwCgYIKwYBBQUHAwQwCgYIKwYBBQUH
# AwEwCgYIKwYBBQUHAwMwCgYIKwYBBQUHAwIwHQYDVR0OBBYEFNa9JauS6LePNjb5
# dFYTlHdVHwezMB8GA1UdIwQYMBaAFEJVytywLIZfaA5FfWDnLTlELaaQMIIBTQYD
# VR0fBIIBRDCCAUAwggE8oIIBOKCCATSGUmh0dHA6Ly9zdWJjYWRldmljZXMucHZh
# LnNvenZlcnMuYXQvUGVuc2lvbnN2ZXJzaWNoZXJ1bmdzYW5zdGFsdCUyMFN1YkNB
# RGV2aWNlcy5jcmyGgd1sZGFwOi8vL0NOPVBlbnNpb25zdmVyc2ljaGVydW5nc2Fu
# c3RhbHQlMjBTdWJDQURldmljZXMsQ049U3ViQ0FEZXZpY2VzLENOPUNEUCxDTj1Q
# dWJsaWMlMjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0
# aW9uLERDPXB2YSxEQz1zb3p2ZXJzLERDPWF0P2NlcnRpZmljYXRlUmV2b2NhdGlv
# bkxpc3Q/YmFzZT9vYmplY3RDbGFzcz1jUkxEaXN0cmlidXRpb25Qb2ludDCCAUgG
# CCsGAQUFBwEBBIIBOjCCATYwYQYIKwYBBQUHMAKGVWh0dHA6Ly9zdWJjYWRldmlj
# ZXMucHZhLnNvenZlcnMuYXQvUGVuc2lvbnN2ZXJzaWNoZXJ1bmdzYW5zdGFsdCUy
# MFN1YkNBRGV2aWNlcygxKS5jcnQwgdAGCCsGAQUFBzAChoHDbGRhcDovLy9DTj1Q
# ZW5zaW9uc3ZlcnNpY2hlcnVuZ3NhbnN0YWx0JTIwU3ViQ0FEZXZpY2VzLENOPUFJ
# QSxDTj1QdWJsaWMlMjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25m
# aWd1cmF0aW9uLERDPXB2YSxEQz1zb3p2ZXJzLERDPWF0P2NBQ2VydGlmaWNhdGU/
# YmFzZT9vYmplY3RDbGFzcz1jZXJ0aWZpY2F0aW9uQXV0aG9yaXR5MA0GCSqGSIb3
# DQEBCwUAA4IBAQBeUsrIB8akC+wcKlFx4GEQOeq04BzpY1GiZksm6EdnRYeq0mXt
# wC4Tx0JBl6d9DhRiWB2qUOrCC/cRppimmQPlgQOqP91j7ihMFaxP1hBcrG4HPxdx
# Z+JYz8ajlXGvGx9m6t4FMKpbgVp1My1Utobo87mAHQEszSIzUsVoeg4RwpqmLwRr
# h2ww2Z8QFTa9mqiWTWVUx8UW88cYHsOiQotlBjVi+ExZ0gNHPU2KFPk5ALpVE2zl
# 3C+XUVbPLOYYKuWhpNViEgD6/yMQ8bDot5cKrFx/KCSdiQTX61WkN7uM72AxsRvZ
# julDpqfz907r9gbV16LAjQ5Kig/Jcp3/L0kAMYIFUzCCBU8CAQEwgY0wdjESMBAG
# CgmSJomT8ixkARkWAmF0MRcwFQYKCZImiZPyLGQBGRYHc296dmVyczETMBEGCgmS
# JomT8ixkARkWA3B2YTEyMDAGA1UEAxMpUGVuc2lvbnN2ZXJzaWNoZXJ1bmdzYW5z
# dGFsdCBTdWJDQURldmljZXMCE0oAAI6eyuHJiZ0aHnkAAQAAjp4wCQYFKw4DAhoF
# AKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisG
# AQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcN
# AQkEMRYEFHsADBv8sJcA97cMz7LOm1dBa0cKMA0GCSqGSIb3DQEBAQUABIIBAB6Q
# kFU2xlK4q+P/PZHw8V4A+SayIc5o4bIraRUvYbFGPH3MM3YfvyTcjXCm3+izLfaC
# diEcpTK4zhN/Mlbbu9BHPuVJJr2XMJr3eonPJlXF2IRrU9iNap3e3MqCp9E9RFIE
# MoGGHbksdxoCZOl2rtL6mqTr/cIPQOyKsHe4j1UaM32Pi84lU84CeXhIHjQYgd4s
# Az6SRaXbjTxicCxrKLf/kvuK8g+idpFPRiS38/h3KjZG7S2r2KJt/8iHknWM/o8a
# puPBZRHiF10p53D/aHPz90+epUdA0oHZ4z646iaEAMCo+R2QpWo4XEnIiPL5Mb7m
# NerITnV5HzrkUPtAzYahggMgMIIDHAYJKoZIhvcNAQkGMYIDDTCCAwkCAQEwdzBj
# MQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMT
# MkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5n
# IENBAhAMTWlyS5T6PCpKPSkHgD1aMA0GCWCGSAFlAwQCAQUAoGkwGAYJKoZIhvcN
# AQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMjMwMTAyMTQyOTI5WjAv
# BgkqhkiG9w0BCQQxIgQgjLBMT+QjLJXmdjzG+REVopEWWHKBwa1bCI6yoHIK+WQw
# DQYJKoZIhvcNAQEBBQAEggIALkUT+sfgKUIruGA0eldS22GuK7j1jP0u1wDTG58S
# Deh3Pl1dQDAiyK4/xaI6Jt9iljxYBi57bUE0wfECBXnkVudvImUVegbgQ/mf6B79
# rDNcziee+GhgrSUZPbpE0wCozsnTBHVr0SXOSq5+8b6eN2ENRHl768YR8JthZbXO
# tQdH2FPNAZB5d759PoyCfsIih7/P8qUUz+btPUwaFUnsrPElpcIoMLMfVrGoVEUA
# E4OJ/tkzk6xvnpcxusmELv3x2lb5K8bjzRueIYWQaofYzXgMbbwpGIlPjHfeyg/7
# 8qtklRWdhOMhw2syHHPfge8K+Y5ws6ITcO4VXl/q6j/dwFXlGRXkBsWXF7J+BkhW
# rPMOxHqEVtOz4rTNFM5i7zOmAwuyXMRsvMpbENR22lAPVixiGg7VuvPrtFeG3P8e
# vULs5gU8SNtTD7IFRSRv9AiHQDP7vhsLv1qcr1rx+uEpQFAIH4ngojdBNiwA5JEH
# 2OqfjQbk4E8rq6bql9r25cvQpkOkAGQCv6ZUs0q52hoTOa5PQ3c3ftDmISzNkk/j
# wJRiYTVgLeMuYRGE673yQNCwVhmWu+surxOLegNEh/ARJxVVnLXT8jnMV3xlTIYU
# 30TfRo3+tILsmYTVU11Ujwl5B/K56gGuDEgB86J6EbLWV+eKNcirJLQR/hzsIGS6
# 5I0=
# SIG # End signature block
