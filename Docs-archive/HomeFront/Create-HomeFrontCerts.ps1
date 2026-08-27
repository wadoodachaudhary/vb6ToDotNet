<#
.SYNOPSIS
    Create certificates for an already-installed HomeFrontPOC on a Windows
    Server. NO changes to HomeFrontPOC code, csproj, or source files.

.DESCRIPTION
    Use this when HomeFrontPOC is already deployed on the Windows machine
    and just needs TLS certificates wired up — for the SQL Server
    connection (welcome-screen opt 6) and for the Blazor HTTPS endpoint.

    What this script does:
      1. Creates a "HomeFront Dev Root CA" in LocalMachine\Root
         (reused if already present)
      2. Issues a SQL Server cert under that root (SANs cover localhost,
         127.0.0.1, the computer's NetBIOS name, and the FQDN). Imports
         into LocalMachine\My.
      3. Issues a Blazor cert under that root with the same SANs. Imports
         into LocalMachine\My AND exports a PFX file (Kestrel needs PFX).
      4. Grants the SQL Server service account READ on the SQL cert's
         private key file.
      5. Writes the cert thumbprint + ForceEncryption=1 to the SQL Server
         registry (SuperSocketNetLib).
      6. Restarts the SQL Server service.
      7. Sets MACHINE-scope environment variables so Kestrel finds the
         Blazor PFX on next start:
            ASPNETCORE_Kestrel__Certificates__Default__Path
            ASPNETCORE_Kestrel__Certificates__Default__Password
       (env vars only — no edits to appsettings.json, no edits to
        HomeFrontPOC source files)
      8. Prints verification commands.

    What this script does NOT do (deliberately, per "no code changes"):
      - Does NOT extract a deployment zip
      - Does NOT edit HomeFrontPOC.dll, HomeFrontPOC.csproj, appsettings.json
      - Does NOT install the app, configure IIS sites, or change app pools
      - Does NOT set Database__Password (the SA password) — that's a
        separate concern; this script only handles certs

    After this script runs, the existing HomeFrontPOC's welcome-screen
    opt 6 ("Encrypt DB connection (TLS)") will work without warnings,
    and the Blazor app started fresh will serve HTTPS on the configured
    port using the new cert.

.PARAMETER Hostname
    The FQDN clients use to reach this server. Will be included as the
    primary SAN on both certs. Default: this machine's
    $env:COMPUTERNAME.$env:USERDNSDOMAIN. Override if clients reach the
    server via a different name (load balancer DNS, etc.).

.PARAMETER CertDir
    Where the Blazor PFX file is written (and where intermediate cert
    files land). Default: C:\HomeFront\certs

.PARAMETER PfxPassword
    Password for the exported Blazor PFX. Default: "changeit". Must
    match whatever Kestrel is configured to expect (we set both ends
    in this script so they always match).

.PARAMETER BlazorHttpsPort
    Port Kestrel listens on for HTTPS. Default: 5443. Used to populate
    ASPNETCORE_URLS.

.PARAMETER BlazorHttpPort
    Fallback HTTP port. Default: 5065. Useful for health checks.

.PARAMETER CertValidityYears
    Leaf-cert expiry. Default: 2. Root CA always 10 years.

.PARAMETER SkipKestrelEnvVars
    Switch — if the customer uses IIS (not Kestrel-direct), the Kestrel
    env vars are unnecessary. The IIS HTTPS binding to the Blazor cert
    is done via IIS Manager / appcmd manually.

.PARAMETER SkipSqlServerWiring
    Switch — set this if SQL Server is on a DIFFERENT machine. Run the
    SQL Server-side steps on that other server instead (sections 2.1
    through 2.3 of WindowsServerCertSetup.md).

.EXAMPLE
    # Run from elevated PowerShell on the server hosting HomeFrontPOC + SQL Server:
    .\Create-HomeFrontCerts.ps1

.EXAMPLE
    # SQL Server is on a different machine — only create the Blazor cert + env vars here:
    .\Create-HomeFrontCerts.ps1 -SkipSqlServerWiring

.EXAMPLE
    # Customer wants IIS hosting (not Kestrel) — skip the env vars, they'll do IIS binding manually:
    .\Create-HomeFrontCerts.ps1 -SkipKestrelEnvVars

.NOTES
    Runbook: Docs/WindowsServerCertSetup.md
    Built:   2026-05-25
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [string]$Hostname,
    [string]$CertDir = "C:\HomeFront\certs",
    [string]$PfxPassword = "changeit",
    [int]$BlazorHttpsPort = 5443,
    [int]$BlazorHttpPort = 5065,
    [int]$CertValidityYears = 2,
    [switch]$SkipKestrelEnvVars,
    [switch]$SkipSqlServerWiring
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Default hostname: this machine's FQDN
if (-not $Hostname) {
    $Hostname = if ($env:USERDNSDOMAIN) {
        "$env:COMPUTERNAME.$env:USERDNSDOMAIN"
    } else {
        $env:COMPUTERNAME
    }
}

# Must be admin
$isAdmin = ([Security.Principal.WindowsPrincipal]`
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltinRole]::Administrator)
if (-not $isAdmin) { throw "Must be run as Administrator." }

# Logging helpers
function Write-Step($Message) {
    Write-Host ""
    Write-Host "━━━ $Message ━━━" -ForegroundColor Cyan
}
function Write-OK($Message)    { Write-Host "  ✓ $Message" -ForegroundColor Green }
function Write-Warn2($Message) { Write-Host "  ⚠ $Message" -ForegroundColor Yellow }

# ──────────────────────────────────────────────────────────────────
# Setup
# ──────────────────────────────────────────────────────────────────
Write-Step "0. Setup"
Write-Host "  Hostname:           $Hostname"
Write-Host "  Cert directory:     $CertDir"
Write-Host "  Blazor port (HTTPS):  $BlazorHttpsPort"
Write-Host "  Wire SQL Server:    $(-not $SkipSqlServerWiring)"
Write-Host "  Set Kestrel envs:   $(-not $SkipKestrelEnvVars)"

if (-not (Test-Path $CertDir)) {
    New-Item -ItemType Directory -Path $CertDir -Force | Out-Null
    Write-OK "Created $CertDir"
}

$sqlSans = @($Hostname, $env:COMPUTERNAME, 'localhost', '127.0.0.1') | Select-Object -Unique
$blazorSans = $sqlSans   # share the SAN list — covers app + SQL on same box

# ──────────────────────────────────────────────────────────────────
# 1. Root CA (reused if present, else created)
# ──────────────────────────────────────────────────────────────────
Write-Step "1. Root CA"

$rootCA = Get-ChildItem Cert:\LocalMachine\Root |
    Where-Object { $_.Subject -eq 'CN=HomeFront Dev Root CA' } |
    Select-Object -First 1

if (-not $rootCA) {
    if ($PSCmdlet.ShouldProcess('Cert:\LocalMachine\Root', 'Create + trust HomeFront Dev Root CA')) {
        $rootCA = New-SelfSignedCertificate `
            -Type Custom -KeySpec Signature `
            -Subject "CN=HomeFront Dev Root CA" `
            -KeyExportPolicy Exportable `
            -HashAlgorithm sha256 -KeyLength 4096 `
            -CertStoreLocation "Cert:\LocalMachine\My" `
            -KeyUsageProperty Sign -KeyUsage CertSign `
            -NotAfter (Get-Date).AddYears(10)
        Move-Item -Path "Cert:\LocalMachine\My\$($rootCA.Thumbprint)" `
                  -Destination "Cert:\LocalMachine\Root"
        $rootCA = Get-ChildItem "Cert:\LocalMachine\Root\$($rootCA.Thumbprint)"
        Write-OK "Created + trusted root CA: $($rootCA.Thumbprint)"
    }
} else {
    Write-OK "Reusing existing root CA: $($rootCA.Thumbprint)"
}

# ──────────────────────────────────────────────────────────────────
# 2. SQL Server cert (signed by root)
# ──────────────────────────────────────────────────────────────────
Write-Step "2. SQL Server cert"

$sqlCert = New-SelfSignedCertificate `
    -Type SSLServerAuthentication `
    -Subject "CN=$Hostname" `
    -DnsName $sqlSans `
    -KeyExportPolicy Exportable -HashAlgorithm sha256 -KeyLength 2048 `
    -CertStoreLocation "Cert:\LocalMachine\My" `
    -Signer $rootCA -NotAfter (Get-Date).AddYears($CertValidityYears)

Write-OK "Issued SQL cert: $($sqlCert.Thumbprint)"
Write-Host "    SANs: $($sqlSans -join ', ')"
Write-Host "    Expires: $($sqlCert.NotAfter.ToString('yyyy-MM-dd'))"

# ──────────────────────────────────────────────────────────────────
# 3. Blazor cert (signed by root) + PFX export
# ──────────────────────────────────────────────────────────────────
Write-Step "3. Blazor cert + PFX export"

$blazorCert = New-SelfSignedCertificate `
    -Type SSLServerAuthentication `
    -Subject "CN=$Hostname" `
    -DnsName $blazorSans `
    -KeyExportPolicy Exportable -HashAlgorithm sha256 -KeyLength 2048 `
    -CertStoreLocation "Cert:\LocalMachine\My" `
    -Signer $rootCA -NotAfter (Get-Date).AddYears($CertValidityYears)

$blazorPfxPath = Join-Path $CertDir 'blazor.pfx'
$pfxPwdSecure = ConvertTo-SecureString -String $PfxPassword -Force -AsPlainText
Export-PfxCertificate -Cert $blazorCert -FilePath $blazorPfxPath -Password $pfxPwdSecure | Out-Null

Write-OK "Issued Blazor cert: $($blazorCert.Thumbprint)"
Write-OK "Exported PFX: $blazorPfxPath"

# Restrict PFX file ACL (only Admins + the current user can read it;
# whoever runs the app will read via the env var path, so add SYSTEM
# and Network Service as a safety net for common service identities)
& icacls $blazorPfxPath /inheritance:r 2>&1 | Out-Null
& icacls $blazorPfxPath /grant `
    "BUILTIN\Administrators:R" "NT AUTHORITY\SYSTEM:R" `
    "NT AUTHORITY\NETWORK SERVICE:R" `
    "$env:USERDOMAIN\$env:USERNAME`:R" 2>&1 | Out-Null
Write-OK "Restricted ACL on $blazorPfxPath"

# ──────────────────────────────────────────────────────────────────
# 4. SQL Server wiring (cert ACL + registry + restart)
# ──────────────────────────────────────────────────────────────────
if ($SkipSqlServerWiring) {
    Write-Step "4. Skipped SQL Server wiring (-SkipSqlServerWiring)"
    Write-Warn2 "Run the same wiring on the machine hosting SQL Server:"
    Write-Warn2 "  - icacls grant SQL service account READ on the cert private key"
    Write-Warn2 "  - HKLM\SOFTWARE\Microsoft\Microsoft SQL Server\<instance>\MSSQLServer\SuperSocketNetLib"
    Write-Warn2 "    Certificate = $($sqlCert.Thumbprint.ToLower())"
    Write-Warn2 "    ForceEncryption = 1"
    Write-Warn2 "  - Restart-Service MSSQLSERVER"
} else {
    Write-Step "4. Wiring SQL Server to use the new cert"

    $sqlSvc = Get-Service -Name "MSSQL*" -ErrorAction SilentlyContinue |
              Where-Object { $_.Status -eq 'Running' } | Select-Object -First 1
    if (-not $sqlSvc) {
        Write-Warn2 "No running MSSQL* service found on this machine. Skipping cert wiring."
        Write-Warn2 "Re-run with -SkipSqlServerWiring on this machine and run the wiring on the SQL host."
    } else {
        # 4a. Grant SQL Server's service account READ on the cert private key file
        $svcAccount = (Get-WmiObject Win32_Service -Filter "Name='$($sqlSvc.Name)'").StartName
        Write-OK "SQL Server runs as: $svcAccount"

        $cert = Get-ChildItem "Cert:\LocalMachine\My\$($sqlCert.Thumbprint)"
        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
        if ($rsa.Key.UniqueName) {
            $candidates = @(
                "$env:ProgramData\Microsoft\Crypto\Keys\$($rsa.Key.UniqueName)",
                "$env:ProgramData\Microsoft\Crypto\RSA\MachineKeys\$($rsa.Key.UniqueName)"
            )
            $keyPath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
            if ($keyPath) {
                & icacls $keyPath /grant "${svcAccount}:R" 2>&1 | Out-Null
                Write-OK "Granted READ on $keyPath to $svcAccount"
            } else {
                Write-Warn2 "Could not locate private-key file. SQL Server may fail to start."
                Write-Warn2 "Manually grant $svcAccount READ on the cert key under $env:ProgramData\Microsoft\Crypto\"
            }
        }

        # 4b. Find the instance registry path
        $instanceNames = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL" -ErrorAction SilentlyContinue
        $instanceId = $null
        if ($instanceNames) {
            $instanceId = $instanceNames.PSObject.Properties |
                Where-Object { $_.Name -notlike 'PS*' } |
                Select-Object -First 1 -ExpandProperty Value
        }

        if ($instanceId) {
            $instanceKey = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$instanceId\MSSQLServer\SuperSocketNetLib"
            if (Test-Path $instanceKey) {
                if ($PSCmdlet.ShouldProcess($instanceKey, "Set Certificate + ForceEncryption")) {
                    Set-ItemProperty -Path $instanceKey -Name "Certificate" -Value $sqlCert.Thumbprint.ToLower()
                    Set-ItemProperty -Path $instanceKey -Name "ForceEncryption" -Value 1
                    Write-OK "Wrote cert thumbprint + ForceEncryption=1 to $instanceKey"
                }

                # 4c. Restart SQL Server
                if ($PSCmdlet.ShouldProcess($sqlSvc.Name, "Restart-Service")) {
                    Restart-Service -Name $sqlSvc.Name -Force
                    Write-OK "Restarted $($sqlSvc.Name)"
                }
            } else {
                Write-Warn2 "Registry path not found: $instanceKey"
                Write-Warn2 "Configure manually via SQL Server Configuration Manager → SQL Server Network Configuration → Protocols → Right-click → Properties → Certificate tab → pick the cert."
            }
        } else {
            Write-Warn2 "Could not find SQL Server instance in registry; do the SQL Config Manager step manually."
        }
    }
}

# ──────────────────────────────────────────────────────────────────
# 5. Kestrel env vars (so the existing HomeFrontPOC.dll picks up the cert)
# ──────────────────────────────────────────────────────────────────
if ($SkipKestrelEnvVars) {
    Write-Step "5. Skipped Kestrel env vars (-SkipKestrelEnvVars)"
    Write-Warn2 "For IIS hosting, bind the Blazor cert ($($blazorCert.Thumbprint)) to your site:"
    Write-Warn2 "  Import-Module WebAdministration"
    Write-Warn2 "  New-WebBinding -Name '<SiteName>' -Protocol https -Port 443 -HostHeader $Hostname -SslFlags 1"
    Write-Warn2 "  (Get-WebBinding -Name '<SiteName>' -Protocol https).AddSslCertificate('$($blazorCert.Thumbprint)','My')"
} else {
    Write-Step "5. Setting Kestrel env vars (machine scope)"

    [Environment]::SetEnvironmentVariable("ASPNETCORE_Kestrel__Certificates__Default__Path", $blazorPfxPath, "Machine")
    [Environment]::SetEnvironmentVariable("ASPNETCORE_Kestrel__Certificates__Default__Password", $PfxPassword, "Machine")
    [Environment]::SetEnvironmentVariable("ASPNETCORE_URLS", "https://*:$BlazorHttpsPort;http://*:$BlazorHttpPort", "Machine")
    Write-OK "ASPNETCORE_Kestrel__Certificates__Default__Path = $blazorPfxPath"
    Write-OK "ASPNETCORE_Kestrel__Certificates__Default__Password = ***"
    Write-OK "ASPNETCORE_URLS = https://*:$BlazorHttpsPort;http://*:$BlazorHttpPort"
    Write-Warn2 "Existing PowerShell sessions DON'T see new env vars. Open a NEW shell to start the app."
}

# ──────────────────────────────────────────────────────────────────
# 6. Summary + verification
# ──────────────────────────────────────────────────────────────────
Write-Step "Done"
Write-Host ""
Write-Host "  Root CA thumbprint:    $($rootCA.Thumbprint)" -ForegroundColor White
Write-Host "  SQL cert thumbprint:   $($sqlCert.Thumbprint)" -ForegroundColor White
Write-Host "  Blazor cert thumbprint: $($blazorCert.Thumbprint)" -ForegroundColor White
Write-Host "  Blazor PFX file:       $blazorPfxPath" -ForegroundColor White
Write-Host "  PFX password:          $PfxPassword" -ForegroundColor White
Write-Host ""
Write-Host "Verify (run from a NEW shell so env vars are visible):" -ForegroundColor Cyan
Write-Host ""
Write-Host "  # 1. SQL Server now serving TLS with our cert:" -ForegroundColor Gray
Write-Host "  sqlcmd -S `"$Hostname,1433`" -U sa -P `"<sa-password>`" -N -Q `"SELECT 'OK' AS Result, encrypt_option FROM sys.dm_exec_connections WHERE session_id = @@SPID`""
Write-Host ""
Write-Host "  # 2. Start the existing HomeFrontPOC.dll — it picks up the env vars:" -ForegroundColor Gray
Write-Host "  cd C:\HomeFront\HomeFrontPOC\bin\Debug\net10.0"
Write-Host "  dotnet HomeFrontPOC.dll"
Write-Host ""
Write-Host "  # 3. Browse to:" -ForegroundColor Gray
Write-Host "  https://${Hostname}:$BlazorHttpsPort/login"
Write-Host ""
Write-Host "  Padlock should show NO warning (chain valid via 'HomeFront Dev Root CA')." -ForegroundColor Cyan
Write-Host "  On the welcome screen, CHECK 'Encrypt DB connection (TLS)' (opt 6)." -ForegroundColor Cyan
Write-Host ""
Write-Host "What did NOT change:" -ForegroundColor Gray
Write-Host "  ✗ HomeFrontPOC.dll, HomeFrontPOC.csproj — untouched"
Write-Host "  ✗ appsettings.json — untouched"
Write-Host "  ✗ Any .cs / .razor source file — untouched"
Write-Host ""
