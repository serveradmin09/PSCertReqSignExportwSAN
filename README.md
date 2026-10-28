# PFX Request & Export Tool

A PowerShell WinForms GUI for requesting certificates from an internal Active Directory Certificate Services (ADCS) environment and exporting them in several common formats (PFX, CER, KEY, PEM), optionally packaged into a password-protected ZIP archive.

It walks through the whole lifecycle of a server certificate request:

1. Build a CSR (Certificate Signing Request) from a hostname (CN), optional Subject Alternative Names (DNS/IP), and a chosen certificate template.
2. Submit the CSR to one of up to five configured issuing CAs.
3. Import the issued certificate into the CA server's local certificate store.
4. Export it back to your chosen output folder as PFX (with private key), CER, KEY and PEM/chain files, using OpenSSL to split the PFX apart.
5. Optionally zip the exported files into a single AES-256, password-protected archive using 7-Zip.
6. Show, for the selected certificate template, which policy/usage OIDs it defines — resolved to friendly names via a direct LDAP query against the AD Certificate Templates container.

Two versions of the script are included:

| File | Language | Notes |
|---|---|---|
| `PPFX_Script_GER_v4_01.ps1` | German UI | Original; assumes CA servers running a German-locale Windows OS (matches `certutil` output text accordingly). |
| `PPFX_Script_ENG_v4_01.ps1` | English UI | Same functionality; assumes CA servers running an English-locale Windows OS. |

Pick whichever matches your CA servers' OS language — see **Known limitations** below for why this matters.

## Prerequisites

- **Windows** with PowerShell and the .NET WinForms assemblies (present by default on any current Windows PowerShell 5.1 or PowerShell 7+ with the Windows Desktop runtime).
- **Active Directory Certificate Services (ADCS)**, with:
  - At least one Enterprise Issuing CA reachable via `certutil -dump` from the machine running the script.
  - Certificate templates published on that CA (`certutil -catemplates`).
  - Permission to remotely execute commands on the CA server(s) via `Invoke-Command` (WinRM/PSRemoting) and to read/write to their `C$` admin share — the script creates the CSR and signs it directly on the CA server.
- **OpenSSL for Windows**, installed to the default path: `C:\Program Files\OpenSSL-Win64\bin\openssl.exe` (or the 32-bit default `C:\Program Files (x86)\OpenSSL-Win32\bin\openssl.exe`). Used to split the exported PFX into separate CER/KEY/PEM files.
- **7-Zip** (optional), installed to its default path (`C:\Program Files\7-Zip\7z.exe` or the x86 equivalent) — only required if you want the "Create ZIP archive" checkbox to work. Native PowerShell `Compress-Archive` isn't used because it can't password-protect an archive.
- Read access to Active Directory (LDAP) from the machine running the script, used to look up the OIDs defined on a certificate template.

## Getting started

1. Open the script that matches your CA servers' OS language (see table above) in a text editor.
2. Edit the **CONFIGURATION** block near the top of the file (see below).
3. Run the script on a machine that has network access to your CA server(s) and AD, with a user account that has the necessary CA/remoting/AD permissions.

## Configuration

All environment-specific values live in a single block near the top of the script, between the `CONFIGURATION` / `END CONFIGURATION` comment markers. Nothing else in the script needs to be touched for normal use.

```powershell
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
```

| Variable | Purpose |
|---|---|
| `$domainSuffix` | Appended to the hostname field's placeholder text and used to validate that the entered CN is fully-qualified for your domain. |
| `$applePrefixTemplateName` | If you have a template that requires `host/<fqdn>` as the CSR subject (common for some Apple/SCEP-style enrollment profiles), put its **internal** template name here. Leave `""` to skip this behavior entirely. |
| `$templatesUsing2048` | A list of internal template names that should use a 2048-bit key instead of the default 4096-bit, when the optional 2048-bit keysize control is present and selected. |
| `$scriptOwner` | Cosmetic — shown in a small footer label ("Script by `$scriptOwner`"). |
| `$bannerQuoteText` / `$bannerQuoteAuthor` | Cosmetic — an optional short quote shown next to the OID panel. Leave both empty to hide this entirely. |

CA servers and certificate templates are **not** configured manually: the script discovers up to five reachable issuing CAs via `certutil -dump`, and reads each one's published templates via `certutil -catemplates`.

## Optional cosmetic features

Two optional, self-contained features ship disabled by default and are safe to ignore:

- **Portrait / logo image** — a `PictureBox` in the top-right of the form that can show a custom Base64-encoded PNG. Paste your own Base64 string between the `@'` / `'@` markers of `$portraitBase64` to enable it; leave it empty to keep the form without a logo.

## Known limitations

- **CA server OS language matters.** The script greps `certutil -catemplates` output for a localized string to filter out "access denied" lines (and to strip the trailing success message), and these strings differ by Windows UI language. The German version (`PFX_Script_v4_01.ps1`) matches German certutil output (`"verweigert"`, `"...erfolgreich ausgeführt."`); the English version (`PFX_Script_v4_01_EN.ps1`) matches English certutil output (`"denied"`, `"...completed successfully."`). If your CA servers run a different locale, you'll need to adjust these two patterns yourself.
- **Maximum 5 CA servers / 15 templates per CA.** These are fixed limits in the current implementation (the GUI's CA and template dropdowns are built from a fixed number of variables rather than a dynamic list).
- **Default install paths only.** OpenSSL and 7-Zip are located by checking their default installation paths; a custom install location isn't detected automatically.
- **Requires remoting + admin share access to the CA server(s).** The script creates the CSR locally, moves it to the CA server's `C:\Temp` via its `C$` share, signs it there via `Invoke-Command`, and exports the resulting PFX back the same way. Your environment must allow this.
- **No input sanitization beyond basic checks.** Treat this as an internal admin tool, not something to expose to untrusted users.

## License

No license has been specified for this project yet. Add one (e.g. MIT) before relying on this repository for anything beyond personal/internal use.
