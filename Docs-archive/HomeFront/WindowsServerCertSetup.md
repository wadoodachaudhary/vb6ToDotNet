# HomeFrontPOC — Windows Server certificate setup & deployment runbook

This runbook is the Windows-Server-side counterpart to the macOS dev cert
setup documented in `memory/dev_certificates.md`. Hand this file to a
Claude (or a human) on the target Windows Server and they should be able
to go from a fresh extract of the deployment zip → fully working
HomeFrontPOC with TLS-encrypted DB connection and HTTPS Blazor endpoint
in 30–60 minutes.

---

## TL;DR — for a Claude landing on the Windows Server

### What you're holding

- **A zip file** named `HomeFrontPOC-YYYY-MM-DD-HHMM.zip` — built on the
  dev Mac by the team's "ship POC" routine and transferred to this
  server (typically via SCP, RDP file share, or USB). The zip
  **already contains** the compiled DLLs, the FlexKit nupkg, this
  runbook, and `Deploy-HomeFront.ps1`. You do **not** create the zip
  on Windows; you only consume it.
- **This runbook** (`Docs/WindowsServerCertSetup.md`) — the detailed
  reference; everything below the TL;DR is for when something doesn't
  Just Work.
- **`Docs/Deploy-HomeFront.ps1`** — one-shot installer/updater that
  automates §1 through §7. Run with `-WhatIf` first to see what it
  would do.

### What `-ZipPath` means

The script's `-ZipPath` parameter is the **path to the input zip on this
Windows machine** — wherever someone dropped it (`C:\Downloads\…\.zip`,
`C:\Temp\…\.zip`, etc.). The script extracts it to `-InstallDir`
(default `C:\HomeFront\`). The zip is INPUT; not something the script
creates.

### Two scenarios — pick the one that matches what you find on the server

**Scenario 1 — Fresh install (no `C:\HomeFront\HomeFrontPOC\` exists yet)**

```powershell
# 1. Extract the zip's Docs folder somewhere temporary so you can read the runbook + script
Expand-Archive C:\Downloads\HomeFrontPOC-YYYY-MM-DD-HHMM.zip C:\Temp -Force
cd C:\Temp\HomeFrontPOC\Docs

# 2. Read this runbook's §0 (prereqs) + §7 (decide which password path)
# 3. Run the deploy script with the matching scenario from §10a:
#    Scenario A (in-house dev, SA password):
.\Deploy-HomeFront.ps1 -ZipPath "C:\Downloads\HomeFrontPOC-YYYY-MM-DD-HHMM.zip" -SaPassword "<pwd>"
#    Scenario B (customer prod, Integrated Security, IIS):
.\Deploy-HomeFront.ps1 -ZipPath "C:\Downloads\HomeFrontPOC-YYYY-MM-DD-HHMM.zip" `
    -UseIntegratedSecurity -HostingModel IisOutOfProcess `
    -ExistingSqlCertThumbprint <hex> -ExistingBlazorCertThumbprint <hex> `
    -AppServiceAccount "DOMAIN\AppPool"
```

The script handles cert generation, SQL Server wiring, Blazor host
config, appsettings.json updates, and a smoke test. Total runtime
~30–60 seconds.

**Scenario 2 — Incremental update (HomeFrontPOC is already running on this server)**

Certs are already installed. SQL Server is already configured. The
welcome-screen toggles already work. You just need to drop new code on
top.

```powershell
# 1. Identify what's running the app:
Get-Process -Name dotnet -ErrorAction SilentlyContinue                      # Kestrel-direct?
Get-Website | Where-Object Name -eq 'HomeFrontPOC'                          # IIS?
Get-Service | Where-Object Name -like 'HomeFront*'                          # Windows service?

# 2. Stop whichever one is running:
Stop-Process -Name dotnet -Force                                            # Kestrel-direct
Stop-Website -Name 'HomeFrontPOC'; Stop-WebAppPool -Name 'HomeFrontPOC'     # IIS
Stop-Service HomeFrontPOC                                                   # Service

# 3. Back up the live appsettings.json (in case the new zip's version overwrites
#    a customer-edited connection string or DevPassword):
Copy-Item C:\HomeFront\HomeFrontPOC\bin\Debug\net10.0\appsettings.json `
          C:\HomeFront\appsettings.backup.json

# 4. Extract the new zip OVER the existing install:
Expand-Archive C:\Downloads\HomeFrontPOC-YYYY-MM-DD-HHMM.zip C:\HomeFront -Force

# 5. Compare appsettings.json — if the customer had edited the connection
#    string or Database:DevPassword, merge those edits back from the backup:
code C:\HomeFront\HomeFrontPOC\bin\Debug\net10.0\appsettings.json   # or notepad

# 6. Restart whatever you stopped in step 2.
```

The certs DON'T need re-installing on an update (they're already in the
Windows cert store; the new zip's `local-packages/FlexKit.<X>.nupkg` is
already pre-built and references the same trusted root).

If FlexKit was bumped in the new zip (compare the `local-packages/` nupkg
filename to what was there before), nothing extra is required on the
server side — the new nupkg ships INSIDE the zip and the new
`HomeFrontPOC.dll` was compiled against it.

### Auth modes recap (full detail in §7)

| Path | Connection string contains | Where password lives | Welcome-screen toggle |
|---|---|---|---|
| 7a (prod) | `User Id=sa;` | env var `Database__Password` | CHECKED (default) |
| 7b (in-house) | `User Id=sa;` | appsettings.json `Database:DevPassword` | UNCHECKED |
| 7c (AD-joined) | `Integrated Security=True;` | — (Windows account) | doesn't matter |

If you don't know which path the server is on: open
`C:\HomeFront\HomeFrontPOC\bin\Debug\net10.0\appsettings.json` and look
at `ConnectionStrings:DefaultConnection`. If you see
`Integrated Security=True`, it's 7c. Otherwise check
`[Environment]::GetEnvironmentVariable("Database__Password","Machine")`
— if set, it's 7a. If empty, look for `Database:DevPassword` in
appsettings.json (7b).

### If something fails

- **Read §9** in this file (Common Windows-Server pitfalls) — 8 known
  failure modes with cause + fix
- **Read the script's transcript** — every step writes a green `✓` or
  yellow `⚠`, easy to grep
- **Check Event Viewer → Windows Logs → Application** for SQL Server or
  IIS errors
- **Last-resort rollback**: extract the PREVIOUS deploy zip over
  `C:\HomeFront\` (or `git checkout` if the server's install is a git
  clone)

---

## Detailed reference

The goal: reproduce on Windows Server what we built in dev:
- SQL Server presents a chain-trusted TLS cert → app can connect with
  `Encrypt=true; TrustServerCertificate=false` (welcome screen opt 6 ON)
- Blazor app serves HTTPS with a chain-trusted cert so browsers don't
  warn
- SA password never lives in `appsettings.json` (set via env var, or
  via `Database:DevPassword` for in-house dev — see [§7](#7-password-source-pick-ONE))

---

## 0. Prerequisites — verify these on the target server first

```powershell
# .NET 10 runtime / SDK
dotnet --info | Select-String "Version"           # should be 10.x

# SQL Server presence
Get-Service -Name "MSSQL*"                         # should list MSSQLSERVER (or named instance)

# Are you Administrator?
([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
# must return True — cert work + IIS bindings + SQL Server config need admin
```

Three decision points to lock in before starting:

| Question | Answer |
|---|---|
| **Cert source** | Path A (PowerShell self-signed, simplest) / Path B (mkcert, matches dev) / Path C (existing Enterprise CA via `certreq`) — pick one based on what's already in place |
| **Blazor host** | Kestrel-direct (`dotnet HomeFrontPOC.dll`) / IIS reverse-proxy with ANCM / IIS in-process — depends on customer environment |
| **DB password source** | Env var `Database__Password` (recommended) / `Database:DevPassword` in appsettings.json + UNCHECK welcome-screen toggle (in-house dev only) |

---

## 1. Generate the certificate(s)

Two leaf certs are needed:
- **SQL Server cert** — SANs MUST include every hostname the app uses to
  connect (`localhost`, `127.0.0.1`, NetBIOS name, FQDN). Common Name
  should be the FQDN.
- **Blazor cert** — SANs MUST include every hostname users will hit in
  their browser (FQDN, possibly the public URL).

You can issue ONE cert that covers both if all SANs fit on one cert.
Simplest is two separate certs.

### Path A — PowerShell self-signed (no extra tools, ~5 min)

```powershell
# 1a. Create a private root CA (do this ONCE per server; reuse for future certs)
$rootCA = New-SelfSignedCertificate `
    -Type Custom `
    -KeySpec Signature `
    -Subject "CN=HomeFront Dev Root CA" `
    -KeyExportPolicy Exportable `
    -HashAlgorithm sha256 `
    -KeyLength 4096 `
    -CertStoreLocation "Cert:\LocalMachine\My" `
    -KeyUsageProperty Sign `
    -KeyUsage CertSign `
    -NotAfter (Get-Date).AddYears(10)

# Trust the root CA system-wide
Move-Item -Path "Cert:\LocalMachine\My\$($rootCA.Thumbprint)" `
          -Destination "Cert:\LocalMachine\Root"

# 1b. Issue SQL Server cert signed by the root
$sqlCert = New-SelfSignedCertificate `
    -Type SSLServerAuthentication `
    -Subject "CN=$env:COMPUTERNAME.$env:USERDNSDOMAIN" `
    -DnsName "$env:COMPUTERNAME.$env:USERDNSDOMAIN", "$env:COMPUTERNAME", "localhost", "127.0.0.1" `
    -KeyExportPolicy Exportable `
    -HashAlgorithm sha256 `
    -KeyLength 2048 `
    -CertStoreLocation "Cert:\LocalMachine\My" `
    -Signer $rootCA `
    -NotAfter (Get-Date).AddYears(2)

# 1c. Issue Blazor cert signed by the root (same SANs if app + DB share a hostname)
$blazorCert = New-SelfSignedCertificate `
    -Type SSLServerAuthentication `
    -Subject "CN=$env:COMPUTERNAME.$env:USERDNSDOMAIN" `
    -DnsName "$env:COMPUTERNAME.$env:USERDNSDOMAIN", "localhost", "127.0.0.1" `
    -KeyExportPolicy Exportable `
    -HashAlgorithm sha256 `
    -KeyLength 2048 `
    -CertStoreLocation "Cert:\LocalMachine\My" `
    -Signer $rootCA `
    -NotAfter (Get-Date).AddYears(2)

# 1d. Export the Blazor cert to PFX (Kestrel needs PFX)
$pfxPwd = ConvertTo-SecureString -String "changeit" -Force -AsPlainText
Export-PfxCertificate -Cert $blazorCert `
    -FilePath "C:\HomeFront\certs\blazor.pfx" `
    -Password $pfxPwd

Write-Host "SQL cert thumbprint:    $($sqlCert.Thumbprint)"
Write-Host "Blazor cert thumbprint: $($blazorCert.Thumbprint)"
Write-Host "Root CA thumbprint:     $($rootCA.Thumbprint)"
```

**Save those thumbprints — you'll need them for the SQL Config Manager
step (§2.3) and the Kestrel config (§3).**

### Path B — mkcert on Windows (matches dev exactly)

```powershell
# Install scoop (or use Chocolatey)
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force
iwr -useb get.scoop.sh | iex
scoop install mkcert

# Create + install root CA into LocalMachine\Root
mkcert -install

# Issue certs
mkdir C:\HomeFront\certs
cd C:\HomeFront\certs
mkcert -cert-file sqlserver.crt -key-file sqlserver.key `
       "$env:COMPUTERNAME.$env:USERDNSDOMAIN" $env:COMPUTERNAME localhost 127.0.0.1
mkcert -pkcs12 -p12-file blazor.pfx `
       "$env:COMPUTERNAME.$env:USERDNSDOMAIN" $env:COMPUTERNAME localhost 127.0.0.1

# mkcert produces PEM cert + PEM key for SQL Server. SQL Server on
# Windows wants the cert IN the Windows cert store (not file-based like
# our Docker setup), so import via PowerShell:
$certPwd = ConvertTo-SecureString -String "" -Force -AsPlainText
# Combine PEM cert + key into a PFX, then import:
openssl pkcs12 -export -out sqlserver.pfx -inkey sqlserver.key -in sqlserver.crt -password pass:changeit
$sqlCert = Import-PfxCertificate -FilePath C:\HomeFront\certs\sqlserver.pfx `
    -CertStoreLocation Cert:\LocalMachine\My `
    -Password (ConvertTo-SecureString "changeit" -AsPlainText -Force) `
    -Exportable
Write-Host "SQL cert thumbprint: $($sqlCert.Thumbprint)"

# Blazor cert: import too (so we can bind by thumbprint if using IIS)
$blazorCert = Import-PfxCertificate -FilePath C:\HomeFront\certs\blazor.pfx `
    -CertStoreLocation Cert:\LocalMachine\My `
    -Password (ConvertTo-SecureString "changeit" -AsPlainText -Force) `
    -Exportable
Write-Host "Blazor cert thumbprint: $($blazorCert.Thumbprint)"
```

### Path C — Existing Enterprise CA (Active Directory Certificate Services)

If the customer has AD CS, request a cert from their CA via `certreq.exe`
or the MMC Certificates snap-in (Computer Account → Personal →
Certificates → Request New Certificate → pick the "Web Server"
template). Make sure the resulting cert's SANs cover everything in §1
above. After enrollment the cert is automatically in
`Cert:\LocalMachine\My`; skip to §2.

---

## 2. Wire the SQL Server cert

Different from Linux/Docker (which uses `mssql.conf`). On Windows SQL
Server reads its TLS cert from the **Windows certificate store** and
is configured via **SQL Server Configuration Manager** or registry.

### 2.1 Grant the SQL Server service account READ access to the cert's private key

This is the #1 cause of "SQL Server failed to start after TLS config".

```powershell
# Find the SQL Server service account (usually NT Service\MSSQLSERVER for
# default instance, or domain account if customer changed it)
$svcAccount = (Get-WmiObject Win32_Service -Filter "Name='MSSQLSERVER'").StartName
Write-Host "SQL Server runs as: $svcAccount"

# Locate the cert's private key file on disk
$cert = Get-ChildItem "Cert:\LocalMachine\My\$($sqlCert.Thumbprint)"
$rsaCert = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
$keyFile = $rsaCert.Key.UniqueName
$keyPath = "$env:ProgramData\Microsoft\Crypto\Keys\$keyFile"
# (Older CSPs: $env:ProgramData\Microsoft\Crypto\RSA\MachineKeys\$keyFile)

# Grant the service account READ on that key file
icacls $keyPath /grant "${svcAccount}:R"
```

### 2.2 Tell SQL Server which cert to use

Easiest: SQL Server Configuration Manager →
**SQL Server Network Configuration → Protocols for MSSQLSERVER →
Right-click → Properties → Certificate tab → pick the cert by friendly
name → Apply**.

Programmatic equivalent (registry — for use in a script):

```powershell
$instanceKey = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQLServer\SuperSocketNetLib"
#                 ^^^^^^                                       ^^^^^^^^^^^^^^^^^^^^^
#                 16 = SQL Server 2022; bump to 17 for next version. Use the actual
#                 instance ID; check HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\InstanceNames\SQL
Set-ItemProperty -Path $instanceKey -Name "Certificate" -Value $sqlCert.Thumbprint.ToLower()
Set-ItemProperty -Path $instanceKey -Name "ForceEncryption" -Value 1
```

`ForceEncryption=1` is the equivalent of the dev `forceencryption=1`
in mssql.conf — server requires TLS on every connection.

### 2.3 Restart SQL Server

```powershell
Restart-Service -Name MSSQLSERVER -Force
```

If the service fails to start, the most likely cause is private-key
permissions (see §2.1). Check the Windows Event Log:
`Get-EventLog -LogName Application -Source MSSQLServer -Newest 20`.

---

## 3. Wire the Blazor / Kestrel cert

Three deployment models — pick the one your customer wants.

### 3a. Kestrel-direct (simplest — `dotnet HomeFrontPOC.dll`)

This matches the dev launchSettings.json `https` profile. The cert is
provided to Kestrel via environment variables OR
`appsettings.{Environment}.json`.

**Option 1 — env vars (recommended for one-off / scripted starts):**

```powershell
# In the same shell session you'll run the app from:
$env:ASPNETCORE_Kestrel__Certificates__Default__Path = "C:\HomeFront\certs\blazor.pfx"
$env:ASPNETCORE_Kestrel__Certificates__Default__Password = "changeit"
$env:ASPNETCORE_URLS = "https://*:5443;http://*:5065"
$env:Database__Password = "<the-SA-password>"        # see §7
cd C:\HomeFront\HomeFrontPOC\bin\Debug\net10.0
dotnet HomeFrontPOC.dll
```

**Option 2 — `appsettings.Production.json` (recommended for unattended start as a Windows service):**

Drop this file next to `HomeFrontPOC.dll`:

```json
{
  "Kestrel": {
    "Endpoints": {
      "Https": { "Url": "https://*:5443" },
      "Http":  { "Url": "http://*:5065" }
    },
    "Certificates": {
      "Default": {
        "Path": "C:\\HomeFront\\certs\\blazor.pfx",
        "Password": "changeit"
      }
    }
  },
  "Database": {
    "Password": "<the-SA-password>"
  }
}
```

Path is double-backslash because JSON. **Restrict file ACLs** so only
the service account can read this file (since it contains the SA
password):

```powershell
icacls "appsettings.Production.json" /inheritance:r
icacls "appsettings.Production.json" /grant "${svcAccount}:R"
icacls "appsettings.Production.json" /grant "BUILTIN\Administrators:F"
```

### 3b. IIS reverse-proxy with ANCM (Out-of-Process)

When IIS terminates HTTPS and proxies plaintext HTTP to Kestrel on a
loopback port:

```powershell
# 1. Install ASP.NET Core Hosting Bundle (one-time per server):
#    https://dotnet.microsoft.com/download/dotnet/10.0 → "Hosting Bundle"

# 2. Create the IIS site and bind the Blazor cert by thumbprint
Import-Module WebAdministration
New-WebSite -Name "HomeFrontPOC" `
    -PhysicalPath "C:\HomeFront\HomeFrontPOC\bin\Debug\net10.0" `
    -Port 80 -HostHeader "$env:COMPUTERNAME.$env:USERDNSDOMAIN"

# 3. Add HTTPS binding using the cert from §1
New-WebBinding -Name "HomeFrontPOC" -Protocol https -Port 443 `
    -HostHeader "$env:COMPUTERNAME.$env:USERDNSDOMAIN" `
    -SslFlags 1                              # SNI on

$binding = Get-WebBinding -Name "HomeFrontPOC" -Protocol https
$binding.AddSslCertificate($blazorCert.Thumbprint, "My")

# 4. App pool needs No Managed Code (Kestrel handles .NET)
Set-ItemProperty IIS:\AppPools\HomeFrontPOC -Name managedRuntimeVersion -Value ""

# 5. Set the env var on the app pool so the app picks up the DB password
$appPool = Get-Item IIS:\AppPools\HomeFrontPOC
$appPool | New-ItemProperty -Name "environmentVariables" -Value @{
  "Database__Password" = "<the-SA-password>"
  "ASPNETCORE_ENVIRONMENT" = "Production"
} -Force
```

The Kestrel-cert config (§3a) is NOT needed when IIS terminates TLS.

### 3c. IIS in-process (ANCM hosting model = InProcess)

Same as 3b but in `web.config` set `<aspNetCore hostingModel="InProcess">`.
Kestrel doesn't run; IIS hosts .NET in-process. All TLS handled by IIS
binding from §3b step 3.

---

## 4. Update connection string in appsettings.json

The shipped `appsettings.json` carries a placeholder DB-less connection
string. Update the `Server=` to point at the right SQL Server hostname
(typically the FQDN if the SQL cert was issued for that):

```json
"ConnectionStrings": {
  "DefaultConnection": "Server=YOUR-SQL-HOST.fqdn,1433;Database=HOMEFRONTSQL;User Id=sa;TrustServerCertificate=true;"
}
```

Notes:
- The `Server=` value MUST be one of the SANs on the SQL Server cert.
  If it isn't, SqlClient with `TrustServerCertificate=false` will
  reject the connection with `host name in certificate is incorrect`.
- `TrustServerCertificate=true` is the **safe fallback**. Even if opt 6
  toggle never gets flipped ON, encryption still happens (server forces
  it via `ForceEncryption=1` from §2.2) — the trust setting just
  controls whether the chain is validated. Leave `=true` here; opt 6
  flips it to `false` at runtime to require strict validation.

---

## 5. Verify SQL Server TLS is working

```powershell
# From a different machine (or another login on this server):
sqlcmd -S "$env:COMPUTERNAME.$env:USERDNSDOMAIN,1433" -U sa -P "<password>" `
       -N -Q "SELECT 'Strict TLS OK' AS Result, encrypt_option FROM sys.dm_exec_connections WHERE session_id = @@SPID"
```

- `-N` = encrypt (no flag value)
- `-N` WITHOUT `-C` = strict cert validation
- Expected: returns `Strict TLS OK | TRUE`

If you get `Login failed for user 'sa'`: SA password mismatch.
If you get `host name in certificate is incorrect`: SAN mismatch — re-issue
the cert with the correct hostname.
If you get `certificate chain was issued by an authority that is not trusted`:
root CA isn't in `Cert:\LocalMachine\Root` — re-run §1's "Move-Item to Root"
step (Path A) or `mkcert -install` (Path B).

---

## 6. Verify Blazor HTTPS is working

```powershell
# From the server itself
Invoke-WebRequest -Uri "https://$env:COMPUTERNAME.$env:USERDNSDOMAIN:5443" `
                  -UseBasicParsing | Select-Object StatusCode
```

Should return 200. From a workstation, browse to the same URL and
verify the padlock shows "Issued by: HomeFront Dev Root CA" (Path A) or
"Issued by: mkcert..." (Path B) or your enterprise CA (Path C).

---

## 7. Password source — pick ONE

### Option 7a — env var `Database__Password` (recommended for prod)

Set BEFORE starting the app:

```powershell
[Environment]::SetEnvironmentVariable("Database__Password", "<password>", "Machine")
# Or for the current shell only:
$env:Database__Password = "<password>"
```

Note the **double underscore** — `Database_Password` (single) silently
fails. Verify with `[Environment]::GetEnvironmentVariable("Database__Password","Machine")`.

The welcome-screen **"Get DB Password from System Secrets"** toggle
stays CHECKED (default). App reads from the env var.

### Option 7b — `Database:DevPassword` in appsettings.json (in-house dev / pre-prod only)

Edit `appsettings.json` on the server:

```json
"Database": {
  "DevPassword": "<password>"
}
```

On the welcome screen, **UNCHECK** "Get DB Password from System Secrets".
App reads from this key. **Restrict file ACLs**:

```powershell
icacls "appsettings.json" /inheritance:r
icacls "appsettings.json" /grant "${svcAccount}:R" "BUILTIN\Administrators:F"
```

**Never use 7b on a server reachable from the public internet** — the
password lives in plaintext in a file that ships with the deploy.

### Option 7c — Integrated Security (Windows Authentication) — NO password anywhere

If the customer prefers Windows-account-based auth to SQL Server (very
common in AD-joined environments), use a connection string with no
password at all:

```json
"ConnectionStrings": {
  "DefaultConnection": "Server=YOUR-SQL-HOST.fqdn,1433;Database=HOMEFRONTSQL;Integrated Security=True;Encrypt=False;TrustServerCertificate=true;"
}
```

Key differences from 7a / 7b:
- **No `User Id=`** clause — Windows identity replaces SQL login
- **No `Password=`** clause — there isn't one
- `Integrated Security=True` (synonyms: `Trusted_Connection=True`,
  `Trusted_Connection=Yes`) is the trigger

**Welcome-screen toggle is a no-op in this mode.** The "Get DB Password
from System Secrets" checkbox state does NOT matter — the app detects
`Integrated Security=True` and skips both the secret-store lookup AND
the appsettings `Database:DevPassword` lookup. Leave the toggle in
whatever state it was (default CHECKED is fine).

The Windows account that the app runs under MUST be configured as a
SQL Server login with appropriate database permissions:

```sql
-- Run AS sysadmin in SSMS against the target SQL Server:
CREATE LOGIN [DOMAIN\AppServiceAccount] FROM WINDOWS;
USE HOMEFRONTSQL;
CREATE USER [DOMAIN\AppServiceAccount] FOR LOGIN [DOMAIN\AppServiceAccount];
EXEC sp_addrolemember 'db_owner', 'DOMAIN\AppServiceAccount';
-- Or instead of db_owner, grant more granular permissions as your DBA prescribes.
```

Identify the Windows account the app actually runs as:
- **Kestrel-direct (§3a)**: the user account that runs `dotnet HomeFrontPOC.dll`. If running as a Windows service, that's the service's "Log on as" account.
- **IIS (§3b/c)**: the app pool identity. By default `IIS AppPool\HomeFrontPOC` (a virtual account); often changed to a domain service account for cross-machine SQL access.

```powershell
# Quick verify — from a PowerShell that the service account can run:
sqlcmd -S "YOUR-SQL-HOST,1433" -E -Q "SELECT SUSER_NAME() AS CurrentLogin, ORIGINAL_LOGIN() AS OriginalLogin, DB_NAME() AS CurrentDb"
# -E means trusted connection (Integrated Security).
# Result should show the Windows account, NOT 'sa'.
```

If the result is `Login failed for user '...'`: the Windows account
isn't in SQL Server's login list — go back and run the `CREATE LOGIN`
above.

**This is the most secure option** for AD-joined Windows Server
deployments: no password in any file, no password in any env var, no
password in any secret store. Auth is delegated entirely to Windows /
Kerberos. Recommended whenever the customer's environment allows.

---

## 8. End-to-end smoke test (the customer-side equivalent)

1. Start the app per §3
2. Browse to `https://<server>:5443/login`
3. Browser shows no warning (cert chain valid)
4. On welcome screen, **CHECK "Encrypt DB connection (TLS)"** (opt 6)
5. Verify "Get DB Password from System Secrets" matches your password
   source (CHECKED for 7a, UNCHECKED for 7b)
6. Pick a connection, log in
7. If you reach `/main`: full chain works (browser TLS → app → SQL TLS
   with strict cert validation)

---

## 9. Common Windows-Server pitfalls

| Symptom | Cause | Fix |
|---|---|---|
| SQL Server fails to start after registry edit | Service account can't read cert private key | §2.1 — `icacls` grant |
| SqlClient: "certificate chain was issued by an authority that is not trusted" | Root CA not in `Cert:\LocalMachine\Root` | §1 (Path A) move-item / `mkcert -install` |
| SqlClient: "host name in certificate is incorrect" | Cert SAN doesn't cover the `Server=` value | Re-issue cert with correct SAN, repeat §1 |
| Browser: "Your connection isn't private" | Browser doesn't see the root CA in OS trust | Path A: root CA in `Cert:\LocalMachine\Root` AND optionally `Cert:\CurrentUser\Root` for per-user browsers |
| App startup: `Database:Password not configured` | Env var wasn't set or used single underscore | §7a — double underscore + `Machine` scope (or current-process scope if launched from that shell) |
| `dotnet HomeFrontPOC.dll` exits immediately | Almost always missing prereq — re-check §0 | `dotnet --info` → must list net10.0 runtime |
| HTTPS port already in use | netsh URL reservation conflict, often from an earlier IIS install | `netsh http show urlacl` → `netsh http delete urlacl url=https://+:5443/` then retry |
| IIS won't bind the cert | Cert isn't in `Cert:\LocalMachine\My` (mkcert sometimes puts it in `CurrentUser`) | `Get-ChildItem Cert:\LocalMachine\My` to verify; re-import with `-CertStoreLocation Cert:\LocalMachine\My` |
| Cert expired | Self-signed certs from Path A have 2y expiry | Re-run §1 with `-NotAfter (Get-Date).AddYears(N)`; restart SQL + Blazor |

---

## 10a. Worked example — `Deploy-HomeFront.ps1`

Everything above is automated in `Docs/Deploy-HomeFront.ps1` (sitting
next to this runbook in the deployment zip). Three real-world
invocations cover most customer scenarios:

### Scenario A — In-house dev box, SQL auth, Kestrel-direct

Simplest setup. PowerShell self-signed cert, SA password from env var,
Kestrel hosting.

```powershell
# Run from an elevated PowerShell:
cd C:\HomeFront\HomeFrontPOC\Docs
.\Deploy-HomeFront.ps1 `
    -ZipPath C:\Downloads\HomeFrontPOC-2026-05-24-1527.zip `
    -SaPassword "Nusr@t7860"
```

Default values produce: `C:\HomeFront\` install dir, self-signed
cert chain, `Database__Password` machine env var, Kestrel on
`https://<hostname>:5443`. Run the app from a new shell to pick up
the env var.

### Scenario B — AD-joined customer server, Integrated Security, IIS, existing cert

Most secure path — no password lives anywhere in app config; auth
delegated to Windows/Kerberos. Customer's CA-issued cert is already
in `Cert:\LocalMachine\My`.

```powershell
.\Deploy-HomeFront.ps1 `
    -ZipPath C:\Downloads\HomeFrontPOC-2026-05-24-1527.zip `
    -UseIntegratedSecurity `
    -CertSource Existing `
    -ExistingSqlCertThumbprint "AB69748FD14F6E8AF024CCD0FDA07C8D8CC85F6E" `
    -ExistingBlazorCertThumbprint "AB69748FD14F6E8AF024CCD0FDA07C8D8CC85F6E" `
    -HostingModel IisOutOfProcess `
    -AppServiceAccount "CONTOSO\HomeFrontAppPool" `
    -SqlServer "sql01.corp.contoso.com"
```

The script will print the exact `CREATE LOGIN [CONTOSO\HomeFrontAppPool]
FROM WINDOWS;` SQL for the DBA to run on the SQL Server. After that,
browse to `https://sql01.corp.contoso.com/login` — IIS terminates
TLS, ANCM reverse-proxies to Kestrel, app talks to SQL via the
Windows token.

### Scenario C — Match macOS dev parity (mkcert), Kestrel, SQL auth

Useful for QA / pre-prod environments that want byte-parity with what
the dev team sees on their Macs.

```powershell
# Install mkcert first (one-time per server):
scoop install mkcert

.\Deploy-HomeFront.ps1 `
    -ZipPath C:\Downloads\HomeFrontPOC-2026-05-24-1527.zip `
    -SaPassword "P@ssw0rd!" `
    -CertSource Mkcert
```

### What the script does

| Step | Action |
|---|---|
| 1 | Verifies admin shell, .NET 10 runtime, SQL Server presence, ANCM (if IIS mode) |
| 2 | Extracts the deployment zip to `-InstallDir` |
| 3 | Issues / imports certs (paths A, B, or C from §1) |
| 4 | If a local SQL Server is running: grants service account READ on the cert private key, writes thumbprint + `ForceEncryption=1` to the registry, restarts SQL |
| 5 | Wires Blazor: sets machine env vars for Kestrel mode, OR creates IIS site + HTTPS binding + app pool for IIS modes |
| 6 | Rewrites `appsettings.json`'s connection string for the chosen auth path, sets `Database__Password` machine env var (or skips for Integrated Security), restricts file ACLs |
| 7 | Optional smoke test — `sqlcmd -N` strict TLS check, `Invoke-WebRequest` against the Blazor HTTPS endpoint (IIS modes only — Kestrel mode prints the command since the app isn't running yet) |
| 8 | Prints a summary with the relevant thumbprints, paths, and next-step commands |

### Flags worth knowing

- `-WhatIf` — show what would happen, change nothing. Always sanity-check
  on first run.
- `-Verbose` — verbose output for each `Set-ItemProperty` etc.
- `-SkipSmokeTest` — useful when SQL Server is on a different host
  (smoke test would fail; that's not a deployment problem).
- `-CertValidityYears <N>` — bump leaf-cert expiry beyond the default 2.

### Re-running

The script is idempotent for most operations:
- Root CA: reused if already present in `Cert:\LocalMachine\Root` with
  Subject `CN=HomeFront Dev Root CA`
- Leaf certs: a fresh pair is issued each run (old ones stay in
  `Cert:\LocalMachine\My` — clean up periodically with
  `Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.NotAfter -lt (Get-Date) } | Remove-Item`)
- IIS site: removed and recreated (use `-WhatIf` if you have other
  bindings on the site you want to preserve)
- `Database__Password` env var: overwritten with the latest `-SaPassword`

### Logging / debugging

PowerShell's transcript covers everything:

```powershell
Start-Transcript -Path C:\HomeFront\deploy-$(Get-Date -Format yyyyMMdd-HHmmss).log
.\Deploy-HomeFront.ps1 ...
Stop-Transcript
```

Each major step in the script writes `━━━ N. Title ━━━` in cyan and
`  ✓ ...` / `  ⚠ ...` lines underneath, so a transcript is grep-able
for failures.

---

## 10. Renewal — when certs expire

Path A and Path B both produce 2-year leaf certs (10-year root for Path A).
~30 days before expiry:

1. Re-run the leaf-cert issuance step from §1 (don't re-create the root)
2. Re-grant private-key permissions per §2.1 (the file path changes
   because the thumbprint is new)
3. Update the SQL Server registry's `Certificate` value (§2.2)
4. Re-bind the new cert to the IIS site OR update the Kestrel PFX file
   (§3)
5. Restart `MSSQLSERVER` and the app pool / Blazor process

Optionally script a cert-rotation job that runs annually under a service
account.

---

## Appendix — files to ship to the server

Beyond the deployment zip (which is built per
`memory/prepare_homefrontpoc_deployment.md`), the server needs:

- The **certificate files** (`blazor.pfx`, `sqlserver.pfx`) — generated
  per §1, stored on the server only, never committed
- The **SA password** — communicated out-of-band, set per §7
- This **runbook** (`Docs/WindowsServerCertSetup.md`) — already in the
  deployment zip
- The **PowerShell scripts** in §1, §2, §3 — copy/paste-ready in the
  shell; consider saving the customer-specific values (paths, hostnames,
  thumbprints) into a per-server `deploy.ps1` for repeatable use

If using Path A and migrating to a new server, EXPORT the root CA from
the old server (`Export-PfxCertificate -Cert $rootCA -FilePath rootCA.pfx
-Password $secure`) and IMPORT into the new server's
`Cert:\LocalMachine\Root` — saves re-issuing every leaf cert on the new
host.
