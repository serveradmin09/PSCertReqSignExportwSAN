Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

# --- Prüfen, ob von UNC-Pfad gestartet, falls ja: lokal kopieren und neu starten ---
$exePath = [System.Reflection.Assembly]::GetEntryAssembly().Location

if ($exePath -like "\\*") {
    $localCopy = Join-Path $env:TEMP (Split-Path $exePath -Leaf)
    try {
        Copy-Item -Path $exePath -Destination $localCopy -Force
        Start-Process -FilePath $localCopy -WorkingDirectory $env:TEMP
        exit
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Konnte Programm nicht lokal kopieren: $_", "Fehler")
        exit 1
    }
}
Set-Location -Path $env:TEMP

# ---------------------------------------------------------------------------------------------#
# This script creates a certificate with subject saternative names (DNS), using a CSR          #
# imports it into the Windows certificate storage (my certificates) on your issuing CA-server  #
# exports the public certificate as CER-File,                                                  #
# the public certificate including private key as PFX-File and                                 #
# the private key as KEY-File to your destination directory .                                  #
# ---------------------------------------------------------------------------------------------#

# =================================================================================================
# CONFIGURATION - adjust these values for your own environment before using this script.
# =================================================================================================

# The domain suffix appended to hostnames (CN) and used to validate user input. Example: "contoso.com"
$domainSuffix = "yourdomain.tld"

# The internal name (not display name) of the certificate template that needs the "host/" CN prefix
# workaround (e.g. for Apple ACME/SCEP-style enrollment templates). Leave empty ("") to disable this.
$applePrefixTemplateName = ""

# Internal template names that should use a 2048-bit key instead of 4096-bit, when the (optional)
# 2048-bit keysize option is selected in the GUI.
$templatesUsing2048 = @("")

# Shown in the small footer label at the bottom of the form.
$scriptOwner = "Your Name"

# Optional short quote shown next to the certificate-template OID list (leave both empty to hide it).
$bannerQuoteText = ""
$bannerQuoteAuthor = ""

# =================================================================================================
# END CONFIGURATION
# =================================================================================================

# checking if OpenSSL is installed #
if (!(Test-Path "C:\Program Files\OpenSSL-Win64\bin\openssl.exe") -or (Test-Path "C:\Program Files (x86)\OpenSSL-Win32\bin\openssl.exe")) {
    [System.Windows.MessageBox]::Show('OpenSSL for Windows was not found in the setup-default directory. Please install OpenSSL first!', 'Warning')
} 
else {}


# creating list of CA's - maximum is 5 #
$cadump = certutil -dump | select-string -Pattern 'onfig' 
$caentries = $cadump -split [System.Environment]::NewLine.ToCharArray()
if ($caentries[0]) {
    $caserver1 = [regex]::match($caentries[0], '"([^"]+)"').Groups[1].Value
}
else {}
if ($caentries[1]) {
    $caserver2 = [regex]::match($caentries[1], '"([^"]+)"').Groups[1].Value
}
else {}
if ($caentries[2]) {
    $caserver3 = [regex]::match($caentries[2], '"([^"]+)"').Groups[1].Value
}
else {}
if ($caentries[3]) {
    $caserver4 = [regex]::match($caentries[3], '"([^"]+)"').Groups[1].Value
}
else {}
if ($caentries[4]) {
    $caserver5 = [regex]::match($caentries[4], '"([^"]+)"').Groups[1].Value
}
else {}

if ($cadump.Count -eq 1) { $calist = $caserver1 }
elseif ($cadump.Count -eq 2) { $calist = $caserver1, $caserver2 }
elseif ($cadump.Count -eq 3) { $calist = $caserver1, $caserver2, $caserver3 }
elseif ($cadump.Count -eq 4) { $calist = $caserver1, $caserver2, $caserver3, $caserver4 }
elseif ($cadump.Count -eq 5) { $calist = $caserver1, $caserver2, $caserver3, $caserver4, $caserver5 }
else {[System.Windows.MessageBox]::Show('No CA found. Make sure, your network and your CA server is reachable, then restart this program.', 'Error')}



# implementing Get-Folder function for output-directory picker #

Function Get-Folder($initialDirectory) {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowserDialog.RootFolder = 'Desktop'
    if ($initialDirectory) { $FolderBrowserDialog.SelectedPath = $initialDirectory }
    [void] $FolderBrowserDialog.ShowDialog()
    return $FolderBrowserDialog.SelectedPath
}

# implementing Get-TemplateOids function to list the OIDs defined in a certificate template, resolved to plain text where possible #
# Queries Active Directory directly via LDAP instead of parsing certutil text output - this avoids any
# dependency on certutil's output format or the OS language of the CA server.

Function Get-TemplateOids ($templateName) {
    try {
        $rootDSE = [ADSI]"LDAP://RootDSE"
        $configNC = $rootDSE.configurationNamingContext
    }
    catch {
        return @("Could not contact Active Directory.")
    }

    try {
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.SearchRoot = [ADSI]"LDAP://CN=Certificate Templates,CN=Public Key Services,CN=Services,$configNC"
        $searcher.Filter = "(&(objectClass=pKICertificateTemplate)(cn=$templateName))"
        $searcher.PropertiesToLoad.AddRange(@("pKIExtendedKeyUsage", "msPKI-Certificate-Application-Policy", "msPKI-Certificate-Policy"))
        $result = $searcher.FindOne()
    }
    catch {
        return @("Error querying Active Directory for this template.")
    }

    if (-not $result) {
        return @("Template '$templateName' not found in Active Directory.")
    }

    $foundOids = New-Object System.Collections.Generic.List[string]

    foreach ($propName in @("pkiextendedkeyusage", "msPKI-Certificate-Application-Policy", "msPKI-Certificate-Policy")) {
        if ($result.Properties.Contains($propName)) {
            foreach ($value in $result.Properties[$propName]) {
                $foundOids.Add([string]$value)
            }
        }
        else {}
    }

    $foundOids = $foundOids | Select-Object -Unique

    if (-not $foundOids) {
        return @("No OIDs found in this template.")
    }

    $oidLines = foreach ($oid in $foundOids) {
        try {
            $oidObj = New-Object System.Security.Cryptography.Oid($oid)
            if ($oidObj.FriendlyName -and ($oidObj.FriendlyName -ne $oid)) {
                "$($oidObj.FriendlyName)"
            }
            else {
                "$oid"
            }
        }
        catch {
            "$oid"
        }
    }
    return $oidLines | Sort-Object -Unique
}

# setting template entries for gui #
# reading templates from CA, splitting them and adding them to the variables for drop down menu


if ($cadump.Count -eq 1) {
$templateentries1 = certutil -catemplates -config $caserver1
$templateentries1 = $templateentries1 | select-string -Pattern "verweigert" -notmatch
$templateentries1 = $templateentries1 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries1 = $templateentries1.split(":")
$templateentries1 = $templateentries1[0] + "," + $templateentries1[2] + "," + $templateentries1[4] + "," + $templateentries1[6] + "," + $templateentries1[8] + "," + $templateentries1[10] + "," + $templateentries1[12] + "," + $templateentries1[14] + "," + $templateentries1[16] + "," + $templateentries1[18] + "," + $templateentries1[20] + "," + $templateentries1[22] + "," + $templateentries1[24] + "," + $templateentries1[26] + "," + $templateentries1[28] + "," + $templateentries1[30]
$templates1 = $templateentries1 -split ","
$templates1 = $templates1 -replace "CertUtil" -replace ""
$templates1 = $templates1 | ? { $_ } | Sort-Object -Unique

} elseif ($cadump.Count -eq 2) {
$templateentries1 = certutil -catemplates -config $caserver1
$templateentries1 = $templateentries1 | select-string -Pattern "verweigert" -notmatch
$templateentries1 = $templateentries1 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries1 = $templateentries1.split(":")
$templateentries1 = $templateentries1[0] + "," + $templateentries1[2] + "," + $templateentries1[4] + "," + $templateentries1[6] + "," + $templateentries1[8] + "," + $templateentries1[10] + "," + $templateentries1[12] + "," + $templateentries1[14] + "," + $templateentries1[16] + "," + $templateentries1[18] + "," + $templateentries1[20] + "," + $templateentries1[22] + "," + $templateentries1[24] + "," + $templateentries1[26] + "," + $templateentries1[28] + "," + $templateentries1[30]
$templates1 = $templateentries1 -split ","
$templates1 = $templates1 -replace "CertUtil" -replace ""
$templates1 = $templates1 | ? { $_ } | Sort-Object -Unique

$templateentries2 = certutil -catemplates -config $caserver2
$templateentries2 = $templateentries2 | select-string -Pattern "verweigert" -notmatch
$templateentries2 = $templateentries2 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries2 = $templateentries2.split(":")
$templateentries2 = $templateentries2[0] + "," + $templateentries2[2] + "," + $templateentries2[4] + "," + $templateentries2[6] + "," + $templateentries2[8] + "," + $templateentries2[10] + "," + $templateentries2[12] + "," + $templateentries2[14] + "," + $templateentries2[16] + "," + $templateentries2[18] + "," + $templateentries2[20] + "," + $templateentries2[22] + "," + $templateentries2[24] + "," + $templateentries2[26] + "," + $templateentries2[28] + "," + $templateentries2[30]
$templates2 = $templateentries2 -split ","
$templates2 = $templates2 -replace "CertUtil" -replace ""
$templates2 = $templates2 | ? { $_ } | Sort-Object -Unique

} elseif ($cadump.Count -eq 3) {
$templateentries1 = certutil -catemplates -config $caserver1
$templateentries1 = $templateentries1 | select-string -Pattern "verweigert" -notmatch
$templateentries1 = $templateentries1 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries1 = $templateentries1.split(":")
$templateentries1 = $templateentries1[0] + "," + $templateentries1[2] + "," + $templateentries1[4] + "," + $templateentries1[6] + "," + $templateentries1[8] + "," + $templateentries1[10] + "," + $templateentries1[12] + "," + $templateentries1[14] + "," + $templateentries1[16] + "," + $templateentries1[18] + "," + $templateentries1[20] + "," + $templateentries1[22] + "," + $templateentries1[24] + "," + $templateentries1[26] + "," + $templateentries1[28] + "," + $templateentries1[30]
$templates1 = $templateentries1 -split ","
$templates1 = $templates1 -replace "CertUtil" -replace ""
$templates1 = $templates1 | ? { $_ } | Sort-Object -Unique

$templateentries2 = certutil -catemplates -config $caserver2
$templateentries2 = $templateentries2 | select-string -Pattern "verweigert" -notmatch
$templateentries2 = $templateentries2 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries2 = $templateentries2.split(":")
$templateentries2 = $templateentries2[0] + "," + $templateentries2[2] + "," + $templateentries2[4] + "," + $templateentries2[6] + "," + $templateentries2[8] + "," + $templateentries2[10] + "," + $templateentries2[12] + "," + $templateentries2[14] + "," + $templateentries2[16] + "," + $templateentries2[18] + "," + $templateentries2[20] + "," + $templateentries2[22] + "," + $templateentries2[24] + "," + $templateentries2[26] + "," + $templateentries2[28] + "," + $templateentries2[30]
$templates2 = $templateentries2 -split ","
$templates2 = $templates2 -replace "CertUtil" -replace ""
$templates2 = $templates2 | ? { $_ } | Sort-Object -Unique

$templateentries3 = certutil -catemplates -config $caserver3
$templateentries3 = $templateentries3 | select-string -Pattern "verweigert" -notmatch
$templateentries3 = $templateentries3 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries3 = $templateentries3.split(":")
$templateentries3 = $templateentries3[0] + "," + $templateentries3[2] + "," + $templateentries3[4] + "," + $templateentries3[6] + "," + $templateentries3[8] + "," + $templateentries3[10] + "," + $templateentries3[12] + "," + $templateentries3[14] + "," + $templateentries3[16] + "," + $templateentries3[18] + "," + $templateentries3[20] + "," + $templateentries3[22] + "," + $templateentries3[24] + "," + $templateentries3[26] + "," + $templateentries3[28] + "," + $templateentries3[30]
$templates3 = $templateentries3 -split ","
$templates3 = $templates3 -replace "CertUtil" -replace ""
$templates3 = $templates3 | ? { $_ } | Sort-Object -Unique

} elseif ($cadump.Count -eq 4) {
$templateentries1 = certutil -catemplates -config $caserver1
$templateentries1 = $templateentries1 | select-string -Pattern "verweigert" -notmatch
$templateentries1 = $templateentries1 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries1 = $templateentries1.split(":")
$templateentries1 = $templateentries1[0] + "," + $templateentries1[2] + "," + $templateentries1[4] + "," + $templateentries1[6] + "," + $templateentries1[8] + "," + $templateentries1[10] + "," + $templateentries1[12] + "," + $templateentries1[14] + "," + $templateentries1[16] + "," + $templateentries1[18] + "," + $templateentries1[20] + "," + $templateentries1[22] + "," + $templateentries1[24] + "," + $templateentries1[26] + "," + $templateentries1[28] + "," + $templateentries1[30]
$templates1 = $templateentries1 -split ","
$templates1 = $templates1 -replace "CertUtil" -replace ""
$templates1 = $templates1 | ? { $_ } | Sort-Object -Unique

$templateentries2 = certutil -catemplates -config $caserver2
$templateentries2 = $templateentries2 | select-string -Pattern "verweigert" -notmatch
$templateentries2 = $templateentries2 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries2 = $templateentries2.split(":")
$templateentries2 = $templateentries2[0] + "," + $templateentries2[2] + "," + $templateentries2[4] + "," + $templateentries2[6] + "," + $templateentries2[8] + "," + $templateentries2[10] + "," + $templateentries2[12] + "," + $templateentries2[14] + "," + $templateentries2[16] + "," + $templateentries2[18] + "," + $templateentries2[20] + "," + $templateentries2[22] + "," + $templateentries2[24] + "," + $templateentries2[26] + "," + $templateentries2[28] + "," + $templateentries2[30]
$templates2 = $templateentries2 -split ","
$templates2 = $templates2 | ? { $_ } | Sort-Object -Unique

$templateentries3 = certutil -catemplates -config $caserver3
$templateentries3 = $templateentries3 | select-string -Pattern "verweigert" -notmatch
$templateentries3 = $templateentries3 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries3 = $templateentries3.split(":")
$templateentries3 = $templateentries3[0] + "," + $templateentries3[2] + "," + $templateentries3[4] + "," + $templateentries3[6] + "," + $templateentries3[8] + "," + $templateentries3[10] + "," + $templateentries3[12] + "," + $templateentries3[14] + "," + $templateentries3[16] + "," + $templateentries3[18] + "," + $templateentries3[20] + "," + $templateentries3[22] + "," + $templateentries3[24] + "," + $templateentries3[26] + "," + $templateentries3[28] + "," + $templateentries3[30]
$templates3 = $templateentries3 -split ","
$templates3 = $templates3 -replace "CertUtil" -replace ""
$templates3 = $templates3 | ? { $_ } | Sort-Object -Unique

$templateentries4 = certutil -catemplates -config $caserver4
$templateentries4 = $templateentries4 | select-string -Pattern "verweigert" -notmatch
$templateentries4 = $templateentries4 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries4 = $templateentries4.split(":")
$templateentries4 = $templateentries4[0] + "," + $templateentries4[2] + "," + $templateentries4[4] + "," + $templateentries4[6] + "," + $templateentries4[8] + "," + $templateentries4[10] + "," + $templateentries4[12] + "," + $templateentries4[14] + "," + $templateentries4[16] + "," + $templateentries4[18] + "," + $templateentries4[20] + "," + $templateentries4[22] + "," + $templateentries4[24] + "," + $templateentries4[26] + "," + $templateentries4[28] + "," + $templateentries4[30]
$templates4 = $templateentries4 -split ","
$templates4 = $templates4 -replace "CertUtil" -replace ""
$templates4 = $templates4 | ? { $_ } | Sort-Object -Unique

} elseif ($cadump.Count -eq 5) {
$templateentries1 = certutil -catemplates -config $caserver1
$templateentries1 = $templateentries1 | select-string -Pattern "verweigert" -notmatch
$templateentries1 = $templateentries1 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries1 = $templateentries1.split(":")
$templateentries1 = $templateentries1[0] + "," + $templateentries1[2] + "," + $templateentries1[4] + "," + $templateentries1[6] + "," + $templateentries1[8] + "," + $templateentries1[10] + "," + $templateentries1[12] + "," + $templateentries1[14] + "," + $templateentries1[16] + "," + $templateentries1[18] + "," + $templateentries1[20] + "," + $templateentries1[22] + "," + $templateentries1[24] + "," + $templateentries1[26] + "," + $templateentries1[28] + "," + $templateentries1[30]
$templates1 = $templateentries1 -split ","
$templates1 = $templates1 -replace "CertUtil" -replace ""
$templates1 = $templates1 | ? { $_ } | Sort-Object -Unique

$templateentries2 = certutil -catemplates -config $caserver2
$templateentries2 = $templateentries2 | select-string -Pattern "verweigert" -notmatch
$templateentries2 = $templateentries2 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries2 = $templateentries2.split(":")
$templateentries2 = $templateentries2[0] + "," + $templateentries2[2] + "," + $templateentries2[4] + "," + $templateentries2[6] + "," + $templateentries2[8] + "," + $templateentries2[10] + "," + $templateentries2[12] + "," + $templateentries2[14] + "," + $templateentries2[16] + "," + $templateentries2[18] + "," + $templateentries2[20] + "," + $templateentries2[22] + "," + $templateentries2[24] + "," + $templateentries2[26] + "," + $templateentries2[28] + "," + $templateentries2[30]
$templates2 = $templateentries2 -split ","
$templates2 = $templates2 -replace "CertUtil" -replace ""
$templates2 = $templates2 | ? { $_ } | Sort-Object -Unique


$templateentries3 = certutil -catemplates -config $caserver3
$templateentries3 = $templateentries3 | select-string -Pattern "verweigert" -notmatch
$templateentries3 = $templateentries3 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries3 = $templateentries3.split(":")
$templateentries3 = $templateentries3[0] + "," + $templateentries3[2] + "," + $templateentries3[4] + "," + $templateentries3[6] + "," + $templateentries3[8] + "," + $templateentries3[10] + "," + $templateentries3[12] + "," + $templateentries3[14] + "," + $templateentries3[16] + "," + $templateentries3[18] + "," + $templateentries3[20] + "," + $templateentries3[22] + "," + $templateentries3[24] + "," + $templateentries3[26] + "," + $templateentries3[28] + "," + $templateentries3[30]
$templates3 = $templateentries3 -split ","
$templates2 = $templates2 -replace "CertUtil" -replace ""
$templates3 = $templates3 | ? { $_ } | Sort-Object -Unique

$templateentries4 = certutil -catemplates -config $caserver4
$templateentries4 = $templateentries4 | select-string -Pattern "verweigert" -notmatch
$templateentries4 = $templateentries4 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries4 = $templateentries4.split(":")
$templateentries4 = $templateentries4[0] + "," + $templateentries4[2] + "," + $templateentries4[4] + "," + $templateentries4[6] + "," + $templateentries4[8] + "," + $templateentries4[10] + "," + $templateentries4[12] + "," + $templateentries4[14] + "," + $templateentries4[16] + "," + $templateentries4[18] + "," + $templateentries4[20] + "," + $templateentries4[22] + "," + $templateentries4[24] + "," + $templateentries4[26] + "," + $templateentries4[28] + "," + $templateentries4[30]
$templates4 = $templateentries4 -split ","
$templates4 = $templates4 -replace "CertUtil" -replace ""
$templates4 = $templates4 | ? { $_ } | Sort-Object -Unique

$templateentries5 = certutil -catemplates -config $caserver5
$templateentries5 = $templateentries5 | select-string -Pattern "verweigert" -notmatch
$templateentries5 = $templateentries5 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries5 = $templateentries5.split(":")
$templateentries5 = $templateentries5[0] + "," + $templateentries5[2] + "," + $templateentries5[4] + "," + $templateentries5[6] + "," + $templateentries5[8] + "," + $templateentries5[10] + "," + $templateentries5[12] + "," + $templateentries5[14] + "," + $templateentries5[16] + "," + $templateentries5[18] + "," + $templateentries5[20] + "," + $templateentries5[22] + "," + $templateentries5[24] + "," + $templateentries5[26] + "," + $templateentries5[28] + "," + $templateentries5[30]
$templates5 = $templateentries5 -split ","
$templates5 = $templates5 -replace "CertUtil" -replace ""
$templates5 = $templates5 | ? { $_ } | Sort-Object -Unique

} elseif ($cadump.Count -eq 6) {
$templateentries1 = certutil -catemplates -config $caserver1
$templateentries1 = $templateentries1 | select-string -Pattern "verweigert" -notmatch
$templateentries1 = $templateentries1 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries1 = $templateentries1.split(":")
$templateentries1 = $templateentries1[0] + "," + $templateentries1[2] + "," + $templateentries1[4] + "," + $templateentries1[6] + "," + $templateentries1[8] + "," + $templateentries1[10] + "," + $templateentries1[12] + "," + $templateentries1[14] + "," + $templateentries1[16] + "," + $templateentries1[18] + "," + $templateentries1[20] + "," + $templateentries1[22] + "," + $templateentries1[24] + "," + $templateentries1[26] + "," + $templateentries1[28] + "," + $templateentries1[30]
$templates1 = $templateentries1 -split ","
$templates1 = $templates1 -replace "CertUtil" -replace ""
$templates1 = $templates1 | ? { $_ } | Sort-Object -Unique

$templateentries2 = certutil -catemplates -config $caserver2
$templateentries2 = $templateentries2 | select-string -Pattern "verweigert" -notmatch
$templateentries2 = $templateentries2 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries2 = $templateentries2.split(":")
$templateentries2 = $templateentries2[0] + "," + $templateentries2[2] + "," + $templateentries2[4] + "," + $templateentries2[6] + "," + $templateentries2[8] + "," + $templateentries2[10] + "," + $templateentries2[12] + "," + $templateentries2[14] + "," + $templateentries2[16] + "," + $templateentries2[18] + "," + $templateentries2[20] + "," + $templateentries2[22] + "," + $templateentries2[24] + "," + $templateentries2[26] + "," + $templateentries2[28] + "," + $templateentries2[30]
$templates2 = $templateentries2 -split ","
$templates2 = $templates2 -replace "CertUtil" -replace ""
$templates2 = $templates2 | ? { $_ } | Sort-Object -Unique


$templateentries3 = certutil -catemplates -config $caserver3
$templateentries3 = $templateentries3 | select-string -Pattern "verweigert" -notmatch
$templateentries3 = $templateentries3 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries3 = $templateentries3.split(":")
$templateentries3 = $templateentries3[0] + "," + $templateentries3[2] + "," + $templateentries3[4] + "," + $templateentries3[6] + "," + $templateentries3[8] + "," + $templateentries3[10] + "," + $templateentries3[12] + "," + $templateentries3[14] + "," + $templateentries3[16] + "," + $templateentries3[18] + "," + $templateentries3[20] + "," + $templateentries3[22] + "," + $templateentries3[24] + "," + $templateentries3[26] + "," + $templateentries3[28] + "," + $templateentries3[30]
$templates3 = $templateentries3 -split ","
$templates2 = $templates2 -replace "CertUtil" -replace ""
$templates3 = $templates3 | ? { $_ } | Sort-Object -Unique

$templateentries4 = certutil -catemplates -config $caserver4
$templateentries4 = $templateentries4 | select-string -Pattern "verweigert" -notmatch
$templateentries4 = $templateentries4 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries4 = $templateentries4.split(":")
$templateentries4 = $templateentries4[0] + "," + $templateentries4[2] + "," + $templateentries4[4] + "," + $templateentries4[6] + "," + $templateentries4[8] + "," + $templateentries4[10] + "," + $templateentries4[12] + "," + $templateentries4[14] + "," + $templateentries4[16] + "," + $templateentries4[18] + "," + $templateentries4[20] + "," + $templateentries4[22] + "," + $templateentries4[24] + "," + $templateentries4[26] + "," + $templateentries4[28] + "," + $templateentries4[30]
$templates4 = $templateentries4 -split ","
$templates4 = $templates4 -replace "CertUtil" -replace ""
$templates4 = $templates4 | ? { $_ } | Sort-Object -Unique

$templateentries5 = certutil -catemplates -config $caserver5
$templateentries5 = $templateentries5 | select-string -Pattern "verweigert" -notmatch
$templateentries5 = $templateentries5 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries5 = $templateentries5.split(":")
$templateentries5 = $templateentries5[0] + "," + $templateentries5[2] + "," + $templateentries5[4] + "," + $templateentries5[6] + "," + $templateentries5[8] + "," + $templateentries5[10] + "," + $templateentries5[12] + "," + $templateentries5[14] + "," + $templateentries5[16] + "," + $templateentries5[18] + "," + $templateentries5[20] + "," + $templateentries5[22] + "," + $templateentries5[24] + "," + $templateentries5[26] + "," + $templateentries5[28] + "," + $templateentries5[30]
$templates5 = $templateentries5 -split ","
$templates5 = $templates5 -replace "CertUtil" -replace ""
$templates5 = $templates5 | ? { $_ } | Sort-Object -Unique

$templateentries6 = certutil -catemplates -config $caserver6
$templateentries6 = $templateentries6 | select-string -Pattern "verweigert" -notmatch
$templateentries6 = $templateentries6 -replace "CertUtil: -CATemplates-Befehl wurde erfolgreich ausgeführt." -replace ""
$templateentries6 = $templateentries6.split(":")
$templateentries6 = $templateentries6[0] + "," + $templateentries6[2] + "," + $templateentries6[4] + "," + $templateentries6[6] + "," + $templateentries6[8] + "," + $templateentries6[10] + "," + $templateentries6[12] + "," + $templateentries6[14] + "," + $templateentries6[16] + "," + $templateentries6[18] + "," + $templateentries6[20] + "," + $templateentries6[22] + "," + $templateentries6[24] + "," + $templateentries6[26] + "," + $templateentries6[28] + "," + $templateentries6[30]
$templates6 = $templateentries6 -split ","
$templates6 = $templates6 -replace "CertUtil" -replace ""
$templates6 = $templates6 | ? { $_ } | Sort-Object -Unique

} else {
[System.Windows.MessageBox]::Show('This program supports a maximum of 5 CA servers only. It seems, there are more.', 'Error')
}

if ($templates1 -cmatch "0x"){
    $templates1 = $null
    }
if ($templates2 -cmatch "0x"){
    $templates2 = $null
    }
if ($templates3 -cmatch "0x"){
    $templates3 = $null
    }
if ($templates4 -cmatch "0x"){
    $templates4 = $null
    }
if ($templates5 -cmatch "0x"){
    $templates5 = $null
    }
if ($templates6 -cmatch "0x"){
    $templates6 = $null
    }
else {}

# GUI script part begins here #

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$objForm = New-Object System.Windows.Forms.Form
$objForm.Backcolor = "#FFF5F5F5"
$objForm.StartPosition = "CenterScreen"
$objForm.Size = New-Object System.Drawing.Size(730, 540)
$objForm.Text = "Create CSR / Export Certificate"

#---infotext/cn---#
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10, 10)
$objLabel.Size = New-Object System.Drawing.Size(400, 20)
$objLabel.Font = [System.Drawing.Font]::new("Arial", 10)
$objLabel.Text = "Hostname eingeben (CN - full qualified!):"
$objForm.Controls.Add($objLabel)

#---text-box hostname---#
$objHostname = New-Object System.Windows.Forms.TextBox
$objHostname.Location = New-Object System.Drawing.Size(11, 30)
$objHostname.Size = New-Object System.Drawing.Size(400, 60)
$objHostname.Text = "<hostname>.$domainSuffix"
$objForm.Controls.Add($objHostname)

#---infotext/issuingCA---#
$objCALabel = New-Object System.Windows.Forms.Label
$objCALabel.Location = New-Object System.Drawing.Size(10, 70)
$objCALabel.Size = New-Object System.Drawing.Size(400, 20)
$objCALabel.Font = [System.Drawing.Font]::new("Arial", 10)
$objCALabel.Text = "Ausstellende CA auswählen:"
# max. listing of CA's currently supported
$objForm.Controls.Add($objCALabel)

#---issuingCA---#
$objCAs = New-Object System.Windows.Forms.Combobox
$objCAs.Location = New-Object System.Drawing.Size(11, 90)
$objCAs.Size = New-Object System.Drawing.Size(400, 60)

# $objCAs.Text = ""
$objCAs.Height = 70
$objForm.Controls.Add($objCAs)
$objForm.Topmost = $True
$objForm.Add_Shown({ $objForm.Activate() })
$objCAs.Items.AddRange($calist) #Issuing CAs will be loaded from variable and listed
$objCAs.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$objCAs.SelectedIndex = "2"

#---infotext/templates---#
$objTemplateLabel = New-Object System.Windows.Forms.Label
$objTemplateLabel.Location = New-Object System.Drawing.Size(10, 130)
$objTemplateLabel.Size = New-Object System.Drawing.Size(400, 20)
$objTemplateLabel.Font = [System.Drawing.Font]::new("Arial", 10)
$objTemplateLabel.Text = "Zertifikatsvorlage auswählen:"
#  max. 15 entries per CA supported
$objForm.Controls.Add($objTemplateLabel)

#---templates-dropdown-menu---#
$objTemplates = New-Object System.Windows.Forms.Combobox
$objTemplates.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$objTemplates.Location = New-Object System.Drawing.Size(11, 150)
$objTemplates.Size = New-Object System.Drawing.Size(400, 60)
$objTemplates.Height = 70
$objForm.Topmost = $True
$objForm.Add_Shown({ $objForm.Activate() })
$objTemplates.Items.AddRange($templates3)
$objForm.Controls.Add($objTemplates)
$objTemplates.SelectedIndex = "0"
# $objTemplates.Text = ""


$objCAs.add_SelectedIndexChanged({
        # Check if the selected item has changed, then load the templates for it
        if ($objCAs.SelectedIndex -eq "0") {
        $objTemplates.Items.Clear()
            if (-not $templates1) {[System.Windows.MessageBox]::Show('No template published or insufficient permissions on the selected CA server.','No template found',0,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
            $objForm.Controls.Remove($objKeysizeLabel)
            $objForm.Controls.Remove($objKeysizeInfo)
            $objForm.Controls.Remove($objKeysize)
            } else {
            $objTemplates.Items.AddRange($templates1)
            $objTemplates.SelectedIndex = "0" 
            }

        } elseif ($objCAs.SelectedIndex -eq "1") {
        $objTemplates.Items.Clear()
            if (-not $templates2) {[System.Windows.MessageBox]::Show('No template published or insufficient permissions on the selected CA server.','No template found',0,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
            $objForm.Controls.Remove($objKeysizeLabel)
            $objForm.Controls.Remove($objKeysizeInfo)
            $objForm.Controls.Remove($objKeysize)
            } else {
        $objTemplates.Items.AddRange($templates2)
        $objTemplates.SelectedIndex = 0
            }
        } elseif  ($objCAs.SelectedIndex -eq "2") {
        $objTemplates.Items.Clear()
            if (-not $templates3) {[System.Windows.MessageBox]::Show('No template published or insufficient permissions on the selected CA server.','No template found',0,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
            $objForm.Controls.Remove($objKeysizeLabel)
            $objForm.Controls.Remove($objKeysizeInfo)
            $objForm.Controls.Remove($objKeysize)
            } else {
        $objTemplates.Items.AddRange($templates3)
        $objTemplates.SelectedIndex = 0
            }
        } elseif  ($objCAs.SelectedIndex -eq "3") {
        $objTemplates.Items.Clear()
            if (-not $templates4) {[System.Windows.MessageBox]::Show('No template published or insufficient permissions on the selected CA server.','No template found',0,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
            $objForm.Controls.Remove($objKeysizeLabel)
            $objForm.Controls.Remove($objKeysizeInfo)
            $objForm.Controls.Remove($objKeysize)
            } else {
        $objTemplates.Items.AddRange($templates4)
        $objTemplates.SelectedIndex = 0
            }
        } elseif  ($objCAs.SelectedIndex -eq "4") {
        $objTemplates.Items.Clear()
            if (-not $templates5) {[System.Windows.MessageBox]::Show('No template published or insufficient permissions on the selected CA server.','No template found',0,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
            $objForm.Controls.Remove($objKeysizeLabel)
            $objForm.Controls.Remove($objKeysizeInfo)
            $objForm.Controls.Remove($objKeysize)
            } else {
        $objTemplates.Items.AddRange($templates5)
        $objTemplates.SelectedIndex = 0 
            }
        } elseif  ($objCAs.SelectedIndex -eq "5") {
        $objTemplates.Items.Clear()
            if (-not $templates6) {[System.Windows.MessageBox]::Show('No template published or insufficient permissions on the selected CA server.','No template found',0,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
            $objForm.Controls.Remove($objKeysizeLabel)
            $objForm.Controls.Remove($objKeysizeInfo)
            $objForm.Controls.Remove($objKeysize)
            } else {
        $objTemplates.Items.AddRange($templates6)
        $objTemplates.SelectedIndex = 0 
            }

        } elseif ($objCAs.Text -eq "") {
        $objTemplates.Items.Clear()

        } else {}
    })




$objTemplates.add_SelectedIndexChanged({
        # Refresh the OID list to reflect the currently selected template
        $objOidList.Text = ""
        if ($objTemplates.Text) {
            $oidResults = Get-TemplateOids -templateName $objTemplates.Text
            $objOidList.Text = ($oidResults -join "`n")
        }
        else {}
    })


#---infotext/SAN-entries---#
$objSANLabel = New-Object System.Windows.Forms.Label
$objSANLabel.Location = New-Object System.Drawing.Size(10, 200)
$objSANLabel.Size = New-Object System.Drawing.Size(400, 20)
$objSANLabel.Font = [System.Drawing.Font]::new("Arial", 10)
$objSANLabel.Text = "Hinzufügen von SAN's (Subject Alternative Names, max: 20):"
$objForm.Controls.Add($objSANLabel)

#---infotext/SAN-entries-help--#
$objSAN2Label = New-Object System.Windows.Forms.Label
$objSAN2Label.Location = New-Object System.Drawing.Size(10, 242)
$objSAN2Label.Size = New-Object System.Drawing.Size(400, 25)
$objSAN2Label.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Italic)
$objSAN2Label.Text = "Beispiel:  server.$domainSuffix,alias.$domainSuffix,server,alias `nTrennzeichen: Komma. (,) ohne Leerzeichen."
$objForm.Controls.Add($objSAN2Label)

#---SAN-entries---#
$objSANs = New-Object System.Windows.Forms.TextBox
$objSANs.Location = New-Object System.Drawing.Size(11, 220)
$objSANs.Size = New-Object System.Drawing.Size(400, 20)
$objSANs.Text = ""
$objForm.Controls.Add($objSANs)

#---infotext/SAN-IP-entries---#
$objIPLabel = New-Object System.Windows.Forms.Label
$objIPLabel.Location = New-Object System.Drawing.Size(10, 285)
$objIPLabel.Size = New-Object System.Drawing.Size(400, 20)
$objIPLabel.Font = [System.Drawing.Font]::new("Arial", 10)
$objIPLabel.Text = "IPv4-Adressen eingeben (optional, max. 4):"
$objForm.Controls.Add($objIPLabel)

#---SAN-IP-entries---#
$objIPs = New-Object System.Windows.Forms.TextBox
$objIPs.Location = New-Object System.Drawing.Size(11, 305)
$objIPs.Size = New-Object System.Drawing.Size(400, 30)
$objIPs.Text = "<1st_IPv4>,<2nd_IPv4>"
$objForm.Controls.Add($objIPs)

#---infotext/own passphrase---#
$objownpwdLabel = New-Object System.Windows.Forms.Label
$objownpwdLabel.Location = New-Object System.Drawing.Size(10, 428)
$objownpwdLabel.Size = New-Object System.Drawing.Size(392, 18)
$objownpwdLabel.Font = [System.Drawing.Font]::new("Arial", 10)
$objownpwdLabel.Text = "Passwort hinzufügen (wird bei Angabe auch für ZIP verwendet):"
$objForm.Controls.Add($objownpwdLabel)

#---own passphrase---#
$objownpwd = New-Object System.Windows.Forms.TextBox
$objownpwd.Location = New-Object System.Drawing.Size(403, 425)
$objownpwd.Size = New-Object System.Drawing.Size(215, 30)
$objownpwd.Font = "Arial"
$objownpwd.PasswordChar = "$(0x25CF|%{[char]$_})"
$objForm.Controls.Add($objownpwd)

#---checkbox for cleartext-password
$checkBox1 = New-Object System.Windows.Forms.CheckBox
$checkBox1.Location = New-Object System.Drawing.Point(625, 429)
$checkBox1.Size = New-Object System.Drawing.Point(15, 15)
$checkBox1.Add_CheckStateChanged({
        if ($checkBox1.Checked) {
            $objownpwd.PasswordChar= 0
        }
        else {
            $objownpwd.PasswordChar = "$(0x25CF|%{[char]$_})"
        }
    })
$objForm.Controls.Add($checkBox1)

#---infotext/checkbox---#
$objcheckboxLabel = New-Object System.Windows.Forms.Label
$objcheckboxLabel.Location = New-Object System.Drawing.Size(640, 429)
$objcheckboxLabel.Size = New-Object System.Drawing.Size(45, 15)
$objcheckboxLabel.Text = "sichtbar"
$objForm.Controls.Add($objcheckboxLabel)

#--Infotext-output-folder--#
$objOutDirLineLeft = New-Object System.Windows.Forms.Panel
$objOutDirLineLeft.Location = New-Object System.Drawing.Size(10, 349)
$objOutDirLineLeft.Size = New-Object System.Drawing.Size(270, 2)
$objOutDirLineLeft.BackColor = [System.Drawing.SystemColors]::ControlDark
$objForm.Controls.Add($objOutDirLineLeft)

$objOutDir = New-Object System.Windows.Forms.Label
$objOutDir.Location = New-Object System.Drawing.Size(290, 340)
$objOutDir.Size = New-Object System.Drawing.Size(130, 20)
$objOutDir.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$objOutDir.Font = [System.Drawing.Font]::new("Arial", 10, [System.Drawing.FontStyle]::Bold)
$objOutDir.Text = "Ausgabe-Optionen"
$objForm.Controls.Add($objOutDir)

$objOutDirLineRight = New-Object System.Windows.Forms.Panel
$objOutDirLineRight.Location = New-Object System.Drawing.Size(425, 349)
$objOutDirLineRight.Size = New-Object System.Drawing.Size(263, 2)
$objOutDirLineRight.BackColor = [System.Drawing.SystemColors]::ControlDark
$objForm.Controls.Add($objOutDirLineRight)

# --- Optional custom branding (portrait image). Not included in this public version - see README.
# Paste your own Base64-encoded PNG between the @' and '@ markers, or delete this whole block
# together with the $picture control below if you don't want a logo/portrait at all.
$portraitBase64 = @'
'@

$picture = $null
if ($portraitBase64) {
    $bytes = [Convert]::FromBase64String($portraitBase64)
    $ms = New-Object System.IO.MemoryStream(,$bytes)
    $img = [System.Drawing.Image]::FromStream($ms)

    $picture = New-Object System.Windows.Forms.PictureBox
    $picture.Location = New-Object Drawing.Point(438,10)
    $picture.Size = New-Object Drawing.Size(250,105)
    $picture.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    #$picture.SizeMode = "Zoom"
    $picture.Image = $img

    $objForm.Controls.Add($picture)
}
else {}

# --- Optional Easter Egg: after 3 clicks on the picture above, a dialog opens with another image. ---
# Not included in this public version. Paste your own Base64-encoded PNG to enable it.
$easterEggBase64 = @'
'@

if ($picture -and $easterEggBase64) {
    $script:pictureClickCount = 0

    $picture.Add_Click({
            $script:pictureClickCount++

            if ($script:pictureClickCount -ge 3) {
                $script:pictureClickCount = 0

                try {
                    $eggBytes = [Convert]::FromBase64String($easterEggBase64)
                    $eggMs = New-Object System.IO.MemoryStream(, $eggBytes)
                    $eggImg = [System.Drawing.Image]::FromStream($eggMs)
                }
                catch {
                    [System.Windows.MessageBox]::Show("Kein gültiges Easter-Egg-Bild hinterlegt.", "Ups", 0, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
                    return
                }

                $eggForm = New-Object System.Windows.Forms.Form
                $eggForm.Text = "?!"
                $eggForm.StartPosition = "CenterParent"
                $eggForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
                $eggForm.MaximizeBox = $false
                $eggForm.MinimizeBox = $false
                $eggForm.ClientSize = New-Object System.Drawing.Size(300, 360)

                $eggPicture = New-Object System.Windows.Forms.PictureBox
                $eggPicture.Location = New-Object System.Drawing.Point(10, 10)
                $eggPicture.Size = New-Object System.Drawing.Size(280, 280)
                $eggPicture.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
                $eggPicture.Image = $eggImg
                $eggForm.Controls.Add($eggPicture)

                $eggOkButton = New-Object System.Windows.Forms.Button
                $eggOkButton.Text = "OK"
                $eggOkButton.Location = New-Object System.Drawing.Point(112, 300)
                $eggOkButton.Size = New-Object System.Drawing.Size(75, 25)
                $eggOkButton.Add_Click({ $eggForm.Close() })
                $eggForm.Controls.Add($eggOkButton)
                $eggForm.AcceptButton = $eggOkButton

                [void] $eggForm.ShowDialog($objForm)
            }
            else {}
        })
}
else {}

#--Output-folder--#
$BrowseButton = New-Object System.Windows.Forms.Button
$BrowseButton.Location = New-Object System.Drawing.Size(11, 373)
$BrowseButton.Size = New-Object System.Drawing.Size(100, 28)
$BrowseButton.Text = "Durchsuchen..."
$BrowseButton.Name = "Browse"

$objDirLabel = New-Object System.Windows.Forms.Label
$objDirLabel.Location = New-Object System.Drawing.Size(125, 371)
$objDirLabel.Size = New-Object System.Drawing.Size(570, 18)
$objDirLabel.Font = [System.Drawing.Font]::new("Arial", 10)
$objDirLabel.Text = "'Durchsuchen...' klicken, um Ausgabepfad anzugeben"
$objForm.Controls.Add($objDirLabel)

# output path infotext
$objoutPathText = New-Object System.Windows.Forms.Label
$objoutPathText.Location = New-Object System.Drawing.Size(125, 387)
$objoutPathText.Size = New-Object System.Drawing.Size(200, 20) 
$objoutPathText.Font = [System.Drawing.Font]::new("Arial", 9)
$objoutPathText.Text = "ist der definierte Ausgabepfad."

#---checkbox for zip
$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox2.Location = New-Object System.Drawing.Point(13, 475)
$checkBox2.Size = New-Object System.Drawing.Point(15, 15)
$checkBox2.Add_CheckStateChanged({
        if ($checkBox2.Checked) {
            
        }
        else {
            
        }
    })
$objForm.Controls.Add($checkBox2)

#---infotext/zip-checkbox---#
$objcheckboxzipLabel = New-Object System.Windows.Forms.Label
$objcheckboxzipLabel.Location = New-Object System.Drawing.Size(28,473.4)
$objcheckboxzipLabel.Size = New-Object System.Drawing.Size(130, 20)
$objcheckboxzipLabel.Font = [System.Drawing.Font]::new("Arial", 10)
$objcheckboxzipLabel.Text = "ZIP-Archiv erstellen"
$objForm.Controls.Add($objcheckboxzipLabel)


#--browse-button--#
$BrowseButton.Add_Click({
        $global:outdir = Get-Folder
        $objDirLabel.Text = $outdir
            if ($outdir) {
            $objDirLabel.Font = [System.Drawing.Font]::new("Arial", 10, [System.Drawing.FontStyle]::Bold)
            $objForm.Controls.Add($objoutPathText)
            } else {
            $objDirLabel.Font = [System.Drawing.Font]::new("Arial", 10)
            $objForm.Controls.Remove($objoutPathText)
            $objDirLabel.Text = "'Durchsuchen...' klicken, um Ausgabepfad anzugeben"
            $objForm.Controls.Add($objDirLabel)
            }
    })

$objForm.Controls.Add($BrowseButton) 


#---Infotext/Banner (optional quote, see CONFIGURATION section at the top)---#
if ($bannerQuoteText) {
    $objOpenSSL = New-Object System.Windows.Forms.Label
    $objOpenSSL.Location = New-Object System.Drawing.Size(438, 118)
    $objOpenSSL.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Italic)
    $objOpenSSL.Size = New-Object System.Drawing.Size(255, 25)
    if ($bannerQuoteAuthor) {
        $objOpenSSL.Text = "`"$bannerQuoteText`"`n                                                        - $bannerQuoteAuthor"
    }
    else {
        $objOpenSSL.Text = "`"$bannerQuoteText`""
    }
    $objForm.Controls.Add($objOpenSSL)
}
else {}

$objOutDirLineSide = New-Object System.Windows.Forms.Panel
$objOutDirLineSide.Location = New-Object System.Drawing.Size(438, 163)
$objOutDirLineSide.Size = New-Object System.Drawing.Size(250, 2)
$objOutDirLineSide.BackColor = [System.Drawing.SystemColors]::ControlDark
$objForm.Controls.Add($objOutDirLineSide)

#---infotext/OIDs---#
$objOidLabel = New-Object System.Windows.Forms.Label
$objOidLabel.Location = New-Object System.Drawing.Size(436, 205)
$objOidLabel.Size = New-Object System.Drawing.Size(250, 15)
$objOidLabel.Font = [System.Drawing.Font]::new("Arial", 9)
$objOidLabel.Text = "OID's der ausgewählten Zertifikats-Vorlage:"
$objForm.Controls.Add($objOidLabel)

#---OIDs (borderless, like the logo)---#
$objOidList = New-Object System.Windows.Forms.Label
$objOidList.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$objOidList.Location = New-Object System.Drawing.Size(438, 225)
$objOidList.Size = New-Object System.Drawing.Size(250, 100)
$objOidList.Font = [System.Drawing.Font]::new("Consolas", 8)
$objOidList.ForeColor = "Black"
$objForm.Controls.Add($objOidList)

#--Create/Export Button--#
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(520, 470)
$OKButton.Size = New-Object System.Drawing.Size(90, 23)
$OKButton.Text = "Erstellen"
$OKButton.Name = "Erstellen"

#---Info---#
$objNFO = New-Object System.Windows.Forms.Label
$objNFO.Location = New-Object System.Drawing.Size(310, 485)
#$objNFO.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$objNFO.Font = [System.Drawing.Font]::new("Verdana", 7, [System.Drawing.FontStyle]::Italic)
$objNFO.Size = New-Object System.Drawing.Size(100, 15)
$objNFO.Text = "Script by $scriptOwner"
$objForm.Controls.Add($objNFO)



#setting keysize in case 2k is selected, in any other case, 4k will be chosen
$OKButton.Add_Click({

        if ($objIPs.Text -eq "<1st_IPv4>,<2nd_IPv4>" -or $objIPs.Text -notlike "*.*") {
            $global:ips = $null
            $global:ip1 = $null
            $global:ip2 = $null
            $global:ip3 = $null
            $global:ip4 = $null
        } else {
            $global:ips = $objIPs.Text.Split(",")
            $global:ip1 = $ips[0]
            $global:ip2 = $ips[1]
            $global:ip3 = $ips[2]
            $global:ip4 = $ips[3]
        }   

        #applying hostname
        $global:name = $objHostname.Text
        $global:shortname = $name.Substring(0, $name.IndexOf('.'))
        
      
        
        #--- CSR data and SANs will be applied here---#
        #--- if the configured "Apple prefix" template is used, "host/" will be added to CN --#

        
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

        # Adding the date extension to the filenames
        $dmy = Get-Date -Format "dd_MM_yyyy"
        $cerb64file = "$shortname" + "_b64" + "$dmy.cer"
        $cerfile = "$shortname" + "_" + "$dmy.cer"
        $pfxfile = "$shortname" + "_" + "$dmy.pfx"
        $keyfile = "$shortname" + "_" + "$dmy.key"
        $inffile = "$shortname" + "_" + "$dmy.inf"
        $csrfile = "$shortname" + "_" + "$dmy.csr"
        $rspfile = "$shortname" + "_" + "$dmy.rsp"
        $pemfile = "$shortname" + "_chain_" + "$dmy.pem"
        $nokeyfile = "$shortname" + "_chain_nokey_" + "$dmy.pem"
        # Maybe this one will be necessary one day
        # $encfile = "$shortname" + "_chain_encrypted_" + "$dmy.pem"


        

        # The script is creating now an CSR (Certificate Signing Request)
        $infPath = Join-Path -Path "$outdir" -ChildPath "$inffile"


        # Checking if user input is useable

        if (-not $outdir) {
        [System.Windows.MessageBox]::Show('Please select an output folder.', 'No output directory selected',0,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
        Return
        } else {}

        if (($name -cmatch "<hostname>") -or ($name -eq "") -or (-not $name)) {
        [System.Windows.MessageBox]::Show('Please type in a hostname (CN).', 'Hostname missing',0,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
        Return
        } else {}

        if ($name -cnotmatch [regex]::Escape($domainSuffix)) {
        [System.Windows.MessageBox]::Show("Please only use a full-qualified hostname.$([System.Environment]::NewLine)Example: hostname.$domainSuffix", "Don't forget the domain suffix!",0,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
        Return
        } else {}

        if (($san1 -cmatch "<") -or ($san1 -cmatch ">")) {
        [System.Windows.MessageBox]::Show("Wrong characters in SAN textfield.$([System.Environment]::NewLine)Please check for mistypes and try again.", 'Error',0,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
        Return
        } else {
        }
                
        if ((-not $san1) -or ($san1 -eq "")) {
        $SANResult = [System.Windows.MessageBox]::Show("No SAN (subject alternative name) has been entered.$([System.Environment]::NewLine)Create certificate without SAN entries?",'Warning',4,[System.Windows.Forms.MessageBoxIcon]::Question)
            if ($SANResult -eq "No") {
            Return
            } else {
            }
        } else {}
		
        if (($objKeysize.SelectedIndex -eq "0") -and ($templatesUsing2048 -contains $objTemplates.Text)) {
            $global:keysize = "2048"
        }
        else {
            $global:keysize = "4096"
        }  

        # Here begins the INF-file content part, which is necessary for the CSR.
        if ($applePrefixTemplateName -and ($objTemplates.Text -eq $applePrefixTemplateName)) {
            $global:infContents =
            @"
[Version]
Signature = "`$Windows NT`$"

[NewRequest]
Subject = "CN=host/$name"
KeySpec = 1
KeyLength = $global:keysize
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
_continue_ = "dns=$san20&"
_continue_ = "IPAddress=$ip1&"
_continue_ = "IPAddress=$ip2&"
_continue_ = "IPAddress=$ip3&"
_continue_ = "IPAddress=$ip4"
"@
        }
        else {
        $global:infContents =
        @"
[Version]
Signature = "`$Windows NT`$"

[NewRequest]
Subject = "CN=$name"
KeySpec = 1
KeyLength = $global:keysize
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
_continue_ = "dns=$san20&"
_continue_ = "IPAddress=$ip1&"
_continue_ = "IPAddress=$ip2&"
_continue_ = "IPAddress=$ip3&"
_continue_ = "IPAddress=$ip4"
"@
        }
        
        $global:ca = $objCAs.Text
        $global:caserver = $ca.Substring(0, $ca.IndexOf('\'))
        $global:template = $objTemplates.Text

        # Checking if INF file already exists
        if (Test-Path "$infpath") {
            Remove-Item "$infpath" -Force
        }
        else {}

        # INF-file will now be saved and CSR will be created.
        $infContents | Out-File -filepath $infPath -force
        Invoke-Command -ComputerName $caserver -ScriptBlock {
            if (!(Test-Path C:\Temp)) {
                New-Item -itemType Directory -Path C:\Temp
            }
            else
            {}
        }
        if (Test-Path "\\$caserver\C$\Temp\$inffile")
        { Remove-Item "\\$caserver\C$\Temp\$inffile" -Force }
        else {}
        
        Get-Item -Path "$infPath" | Move-Item -Destination "\\$caserver\C$\Temp"
        
        # Checking, if other files already exist
        if (Test-Path "\\$caserver\C$\Temp\$csrfile")
        { Remove-Item "\\$caserver\C$\Temp\$csrfile" -Force }
        else {}
        if (Test-Path "\\$caserver\C$\Temp\$rspfile")
        { Remove-Item "\\$caserver\C$\Temp\$rspfile" -Force }
        else {}
        Invoke-Command -ComputerName $caserver -ScriptBlock {
            certreq -new "C:\Temp\$using:inffile" "C:\Temp\$using:csrfile"
        }
        # creating target directory
        $targetdir = "$outdir" + "\" + "$shortname"
        if (Test-Path "$targetdir.old")
        { Remove-Item "$targetdir.old" -Force }
        else {}
        if (Test-Path "$targetdir")
        { Rename-Item -Path "$targetdir" "$targetdir.old" -Force }
        else {}
        New-Item $targetdir -ItemType Directory

        # certificate request (CSR) will be signed by your CA, the received certificate will be saved in the destination directory
        
        Invoke-Command -ComputerName $caserver -ScriptBlock { 
            certreq -submit -config "$using:ca" -attrib certificateTemplate:$using:template "C:\Temp\$using:csrfile" "C:\Temp\$using:cerfile" 
        }

        # Importing the just created certificate to the ca-server under "my certificates".
        # After then, key and cert will be exported as PFX to the destination folder.
        # Using the thumbprint (is unique), ther certificate can be retrieved from the server.
        
        $global:thumbprint = Invoke-Command -ComputerName $caserver -ScriptBlock { Import-Certificate -FilePath "C:\Temp\$using:cerfile" -Confirm:$false -CertStoreLocation "Cert:\LocalMachine\My" | Select-Object Thumbprint }
        $global:thumb = $thumbprint.Thumbprint

        # Passphrase = hostname written backwards. 
        # The Passphrase will now be converted to a secure-string.
        # For OpenSSL, the clear-text password will be used, because OpenSSL doesn't recognize a secure-string.
        # If user typed in an own passphrase, it will be used.

        if (!$objownpwd.Text) {
            $global:pwd = $shortname
            $plainpwd = $pwd[-1.. - $pwd.Length] -join ''
            $global:securepwd = convertto-securestring -String $plainpwd -asplaintext -force
        }
        else {
            $global:ownpwd = $objownpwd.Text
            $global:plainpwd = $ownpwd
            $global:securepwd = convertto-securestring -String $plainpwd -asplaintext -force
            $global:pwd = $ownpwd
        }


        # The PFX file is now being saved and protected with the passphrase.

              
        Invoke-Command -ComputerName $caserver -ScriptBlock {
            Get-ChildItem -Path "Cert:\LocalMachine\my\$using:thumb" | Export-PfxCertificate -FilePath "C:\Temp\$using:pfxfile" -Password $using:securepwd
        }
        Get-Item -Path "\\$caserver\C$\Temp\$pfxfile" | Move-Item -Destination "$outdir\$shortname\$pfxfile"
        Get-Item -Path "\\$caserver\C$\Temp\$cerfile" | Move-Item -Destination "$outdir\$shortname\$cerfile"

        if (Test-Path "C:\Program Files\OpenSSL-Win64\bin\openssl.exe") {
            $openssldir = "C:\Program Files\OpenSSL-Win64\bin\openssl.exe"
        } 
        else {
            $openssldir = "C:\Program Files (x86)\OpenSSL-Win32\bin\openssl.exe"
        }

        

        # OpenSSL splits the certificate and private key so we have two seperate files: The key (.KEY-file) and the certificate (.CER-Datei).
        
        & "$openssldir" pkcs12 -in "$outdir\$shortname\$pfxfile"  -out "$outdir\$shortname\$cerb64file" -nodes -nokeys -passin pass:$plainpwd
        & "$openssldir" pkcs12 -in "$outdir\$shortname\$pfxfile"  -out "$outdir\$shortname\$pfxfile.tmp" -passin pass:$plainpwd -passout pass:$plainpwd
        & "$openssldir" pkey -in "$outdir\$shortname\$pfxfile.tmp" -out "$outdir\$shortname\$keyfile" -passin pass:$plainpwd
        & "$openssldir" pkcs12 -in "$outdir\$shortname\$pfxfile" -out "$outdir\$shortname\$pemfile" -noenc -passin pass:$plainpwd
        & "$openssldir" pkcs12 -in "$outdir\$shortname\$pfxfile" -out "$outdir\$shortname\$nokeyfile" -noenc -nokeys -passin pass:$plainpwd
        & "$openssldir" pkcs12 -in "$outdir\$shortname\$pfxfile" -out "$outdir\$shortname\$pemfile" -noenc -passin pass:$plainpwd
        # Maybe this one will be necessary one day
        # cmd /c "$openssldir" pkcs12 -in "$outdir\$shortname\$pfxfile" -out "$outdir\$shortname\$encfile" -passin pass:$plainpwd
        Remove-Item -Force "$outdir\$shortname\$pfxfile.tmp"

        # If checkBox2 is checked, the exported files will be packed into a password-protected ZIP archive instead of staying loose.
        # 7-Zip is used because native Compress-Archive does not support password protection.
        $zipCreated = $false
        if ($checkBox2.Checked) {

            $zipfile = "$shortname" + "_" + "$dmy" + ".zip"

            $sevenzip = $null
            if (Test-Path "C:\Program Files\7-Zip\7z.exe") {
                $sevenzip = "C:\Program Files\7-Zip\7z.exe"
            }
            elseif (Test-Path "C:\Program Files (x86)\7-Zip\7z.exe") {
                $sevenzip = "C:\Program Files (x86)\7-Zip\7z.exe"
            }
            else {}

            if (-not $sevenzip) {
                [System.Windows.MessageBox]::Show("7-Zip was not found in the setup-default directory. The exported files could not be zipped.$([System.Environment]::NewLine)Please install 7-Zip and try again.", "7-Zip missing",0,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
            }
            else {
                $filesToZip = Get-ChildItem -Path "$outdir\$shortname" -File

                & "$sevenzip" a -tzip -mem=AES256 "-p$plainpwd" "$outdir\$shortname\$zipfile" @($filesToZip.FullName) | Out-Null

                if (Test-Path "$outdir\$shortname\$zipfile") {
                    $filesToZip | Remove-Item -Force
                    $zipCreated = $true
                }
                else {
                    [System.Windows.MessageBox]::Show("Zipping the exported files failed.$([System.Environment]::NewLine)The individual files remain in the output directory.", "Zip creation failed",0,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
                }
            }
        }
        else {}

        if ($zipCreated) {
            [void] [System.Windows.MessageBox]::Show("Zertifikat erstellt und exportiert.`n`nPasswortgeschütztes ZIP-Archiv erstellt.", "Erfolg",0,[System.Windows.Forms.MessageBoxIcon]::Information)
        }
        else {
            [void] [System.Windows.MessageBox]::Show("Zertifikat erstellt und exportiert.`n", "Erfolg",0,[System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
$objForm.Controls.Add($OKButton) 

#--Cancel-Button--#
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(615, 470)
$CancelButton.Size = New-Object System.Drawing.Size(80, 23)
$CancelButton.Text = "Schließen"
$CancelButton.Name = "Schließen"
$CancelButton.DialogResult = "Cancel"
$CancelButton.Add_Click({ $objForm.Close() })
$objForm.Controls.Add($CancelButton) 

# Populate the OID text once at startup for the default template - the initial
# $objTemplates.SelectedIndex assignment above happens before the
# SelectedIndexChanged handler is registered, so it wouldn't fire on its own.
$objOidList.Text = ""
if ($objTemplates.Text) {
    $oidResults = Get-TemplateOids -templateName $objTemplates.Text
    $objOidList.Text = ($oidResults -join "`n")
}
else {}

[void] $objForm.ShowDialog()
