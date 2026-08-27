<#
.SYNOPSIS
    End-to-end Windows Server deployment for HomeFrontPOC.

.DESCRIPTION
    One-shot deploy script that:
      1. Verifies prerequisites (admin, .NET 10, SQL Server presence)
      2. Extracts the deployment zip to -InstallDir
      3. Issues or imports certificates based on -CertSource
      4. Wires the SQL Server cert into the Windows cert store + registry
      5. Wires the Blazor cert into Kestrel or IIS based on -HostingModel
      6. Updates appsettings.json connection string for the chosen
         password source (env var / appsettings.DevPassword / Integrated
         Security)
      7. Optionally runs a smoke test (sqlcmd strict TLS + HTTPS GET)

    Idempotent where possible — re-running on a configured server should
    update mutable bits without re-creating immutable ones.

    See Docs/WindowsServerCertSetup.md for the manual walkthrough this
    script automates.

.PARAMETER ZipPath
    Path to the HomeFrontPOC deployment zip (the artifact produced by
    the macOS-side "prepare HomeFrontPOC for deployment" routine).

.PARAMETER InstallDir
    Root install folder. Default: C:\HomeFront. The zip's HomeFrontPOC/
    contents are extracted under this path so the app ends up at
    C:\HomeFront\HomeFrontPOC\bin\Debug\net10.0\HomeFrontPOC.dll.

.PARAMETER CertDir
    Where issued cert files (sqlserver.pfx, blazor.pfx, rootCA.pfx) are
    written. Default: $InstallDir\certs.

.PARAMETER SqlServer
    SQL Server hostname the app will connect to. Default: FQDN of this
    machine. MUST match a SAN on the SQL cert.

.PARAMETER SqlPort
    SQL Server TCP port. Default: 1433.

.PARAMETER DatabaseName
    Database to put in the connection string. Default: HOMEFRONTSQL.

.PARAMETER UseIntegratedSecurity
    Switch — use Windows Authentication. No SA password needed.
    The connection string becomes
    "Server=...;Database=...;Integrated Security=True;..." and the
    welcome-screen "Get DB Password from System Secrets" toggle becomes
    a no-op. RECOMMENDED for AD-joined customer environments.

.PARAMETER SaPassword
    SA password (string). Required when -UseIntegratedSecurity is NOT
    set. Stored as the Database__Password machine-scope env var (path
    7a in the runbook).

.PARAMETER CertSource
    SelfSigned (PowerShell New-SelfSignedCertificate), Mkcert (requires
    mkcert.exe on PATH), or Existing (use -ExistingSqlCertThumbprint /
    -ExistingBlazorCertThumbprint). Default: SelfSigned.

.PARAMETER ExistingSqlCertThumbprint
    Required when -CertSource Existing.

.PARAMETER ExistingBlazorCertThumbprint
    Required when -CertSource Existing.

.PARAMETER HostingModel
    Kestrel — run as a console process or Windows service via
    `dotnet HomeFrontPOC.dll`. Simplest; matches dev.
    IisOutOfProcess — IIS terminates HTTPS, reverse-proxies HTTP to
    Kestrel on a loopback port. Standard Windows enterprise pattern.
    IisInProcess — IIS hosts .NET in-process (web.config sets
    hostingModel="InProcess"). Slightly faster than OOP.
    Default: Kestrel.

.PARAMETER BlazorHttpsPort
    HTTPS port for Kestrel hosting mode. Default: 5443.

.PARAMETER BlazorHttpPort
    HTTP port for Kestrel hosting mode (fallback / health-check).
    Default: 5065.

.PARAMETER AppServiceAccount
    Windows account the app will run as (Kestrel-via-service or IIS
    AppPool identity). Used to grant private-key ACLs on the SQL cert
    AND, with -UseIntegratedSecurity, to identify which account needs
    a SQL Server login. Default: the current user.

.PARAMETER CertValidityYears
    Cert expiry for self-signed leaves. Default: 2.

.PARAMETER SkipSmokeTest
    Skip the post-deploy verification step (sqlcmd + Invoke-WebRequest).

.EXAMPLE
    # In-house dev box with SQL auth — secret as env var:
    .\Deploy-HomeFront.ps1 `
        -ZipPath C:\Downloads\HomeFrontPOC-2026-05-24-1527.zip `
        -SaPassword "Nusr@t7860"

.EXAMPLE
    # AD-joined customer server with Integrated Security + IIS + corporate CA cert:
    .\Deploy-HomeFront.ps1 `
        -ZipPath C:\Downloads\HomeFrontPOC-2026-05-24-1527.zip `
        -UseIntegratedSecurity `
        -CertSource Existing `
        -ExistingSqlCertThumbprint "AB69748FD14F6E8AF024CCD0FDA07C8D8CC85F6E" `
        -ExistingBlazorCertThumbprint "AB69748FD14F6E8AF024CCD0FDA07C8D8CC85F6E" `
        -HostingModel IisOutOfProcess `
        -AppServiceAccount "CONTOSO\HomeFrontAppPool"

.EXAMPLE
    # Quick Kestrel-direct setup with mkcert (matches macOS dev parity):
    .\Deploy-HomeFront.ps1 `
        -ZipPath C:\Downloads\HomeFrontPOC-2026-05-24-1527.zip `
        -SaPassword "P@ssw0rd!" `
        -CertSource Mkcert

.NOTES
    Runbook: Docs/WindowsServerCertSetup.md
    Memory:  memory/dev_certificates.md, memory/sa_password_secret.md
    Built:   2026-05-24
#>
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({ Test-Path $_ })]
    [string]$ZipPath,

    [string]$InstallDir = "C:\HomeFront",
    [string]$CertDir,
    [string]$SqlServer = "$env:COMPUTERNAME.$env:USERDNSDOMAIN",
    [int]$SqlPort = 1433,
    [string]$DatabaseName = "HOMEFRONTSQL",

    [switch]$UseIntegratedSecurity,
    [string]$SaPassword,

    [ValidateSet('SelfSigned', 'Mkcert', 'Existing')]
    [string]$CertSource = 'SelfSigned',
    [string]$ExistingSqlCertThumbprint,
    [string]$ExistingBlazorCertThumbprint,

    [ValidateSet('Kestrel', 'IisOutOfProcess', 'IisInProcess')]
    [string]$HostingModel = 'Kestrel',
    [int]$BlazorHttpsPort = 5443,
    [int]$BlazorHttpPort = 5065,

    [string]$AppServiceAccount = "$env:USERDOMAIN\$env:USERNAME",
    [int]$CertValidityYears = 2,

    [switch]$SkipSmokeTest
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Default cert dir to InstallDir\certs unless overridden
if (-not $CertDir) { $CertDir = Join-Path $InstallDir 'certs' }

# Parameter validation that PowerShell can't express declaratively
if (-not $UseIntegratedSecurity -and -not $SaPassword) {
    throw "Must specify either -SaPassword or -UseIntegratedSecurity."
}
if ($UseIntegratedSecurity -and $SaPassword) {
    Write-Warning "Both -UseIntegratedSecurity and -SaPassword passed; -SaPassword will be ignored."
    $SaPassword = $null
}
if ($CertSource -eq 'Existing') {
    if (-not $ExistingSqlCertThumbprint -or -not $ExistingBlazorCertThumbprint) {
        throw "-CertSource Existing requires both -ExistingSqlCertThumbprint and -ExistingBlazorCertThumbprint."
    }
}

# ──────────────────────────────────────────────────────────────────
# Logging helpers
# ──────────────────────────────────────────────────────────────────
function Write-Step($Message) {
    Write-Host ""
    Write-Host "━━━ $Message ━━━" -ForegroundColor Cyan
}
function Write-OK($Message)   { Write-Host "  ✓ $Message" -ForegroundColor Green }
function Write-Warn2($Message) { Write-Host "  ⚠ $Message" -ForegroundColor Yellow }

# ──────────────────────────────────────────────────────────────────
# 1. Prerequisites
# ──────────────────────────────────────────────────────────────────
Write-Step "1. Verifying prerequisites"

# Admin?
$isAdmin = ([Security.Principal.WindowsPrincipal]`
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltinRole]::Administrator)
if (-not $isAdmin) { throw "Must be run as Administrator." }
Write-OK "Administrator: yes"

# .NET 10 runtime?
$dotnet = (dotnet --list-runtimes 2>$null) -join "`n"
if ($dotnet -notmatch 'Microsoft\.AspNetCore\.App 10\.') {
    throw ".NET 10 ASP.NET Core runtime not found. Install the Hosting Bundle from https://dotnet.microsoft.com/download/dotnet/10.0"
}
Write-OK ".NET 10 runtime: present"

# SQL Server?
$sqlSvc = Get-Service -Name "MSSQL*" -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Running' }
if (-not $sqlSvc) {
    Write-Warn2 "No running MSSQL* service found. If SQL Server is on a different host, that's expected."
} else {
    Write-OK "SQL Server service: $($sqlSvc[0].Name) ($($sqlSvc[0].Status))"
}

# Hosting Bundle (needed for IIS modes)
if ($HostingModel -ne 'Kestrel') {
    $ancm = Get-WebGlobalModule -Name 'AspNetCoreModuleV2' -ErrorAction SilentlyContinue
    if (-not $ancm) {
        throw "ASP.NET Core Module v2 not installed; required for -HostingModel $HostingModel. Install the Hosting Bundle from dotnet.microsoft.com."
    }
    Write-OK "ASP.NET Core Module v2: present"
}

# ──────────────────────────────────────────────────────────────────
# 2. Extract zip
# ──────────────────────────────────────────────────────────────────
Write-Step "2. Extracting deployment zip to $InstallDir"

if (-not (Test-Path $InstallDir)) {
    New-Item -ItemType Directory -Path $InstallDir | Out-Null
    Write-OK "Created $InstallDir"
}
if (-not (Test-Path $CertDir)) {
    New-Item -ItemType Directory -Path $CertDir | Out-Null
    Write-OK "Created $CertDir"
}

$appDir = Join-Path $InstallDir 'HomeFrontPOC'
if (Test-Path $appDir) {
    Write-Warn2 "$appDir already exists; existing files will be overwritten on conflict."
}
if ($PSCmdlet.ShouldProcess($InstallDir, "Expand-Archive $ZipPath")) {
    Expand-Archive -Path $ZipPath -DestinationPath $InstallDir -Force
    Write-OK "Extracted to $appDir"
}

$dllPath = Join-Path $appDir 'bin\Debug\net10.0\HomeFrontPOC.dll'
if (-not (Test-Path $dllPath)) {
    throw "Expected $dllPath after extraction — zip layout unexpected."
}

# ──────────────────────────────────────────────────────────────────
# 3. Issue / import certificates
# ──────────────────────────────────────────────────────────────────
Write-Step "3. Provisioning certificates ($CertSource)"

$sqlSans = @($SqlServer, $env:COMPUTERNAME, 'localhost', '127.0.0.1') | Select-Object -Unique
$blazorSans = @($SqlServer, $env:COMPUTERNAME, 'localhost', '127.0.0.1') | Select-Object -Unique

$sqlCertThumbprint = $null
$blazorCertThumbprint = $null
$blazorPfxPath = $null
$blazorPfxPassword = 'changeit'

switch ($CertSource) {
    'SelfSigned' {
        # Root CA — reuse if present, else create
        $rootCA = Get-ChildItem Cert:\LocalMachine\Root |
            Where-Object { $_.Subject -eq 'CN=HomeFront Dev Root CA' } |
            Select-Object -First 1
        if (-not $rootCA) {
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
        } else {
            Write-OK "Reusing existing root CA: $($rootCA.Thumbprint)"
        }

        $sqlCert = New-SelfSignedCertificate `
            -Type SSLServerAuthentication `
            -Subject "CN=$SqlServer" `
            -DnsName $sqlSans `
            -KeyExportPolicy Exportable -HashAlgorithm sha256 -KeyLength 2048 `
            -CertStoreLocation "Cert:\LocalMachine\My" `
            -Signer $rootCA -NotAfter (Get-Date).AddYears($CertValidityYears)
        $sqlCertThumbprint = $sqlCert.Thumbprint
        Write-OK "Issued SQL cert: $sqlCertThumbprint (SANs: $($sqlSans -join ', '))"

        $blazorCert = New-SelfSignedCertificate `
            -Type SSLServerAuthentication `
            -Subject "CN=$SqlServer" `
            -DnsName $blazorSans `
            -KeyExportPolicy Exportable -HashAlgorithm sha256 -KeyLength 2048 `
            -CertStoreLocation "Cert:\LocalMachine\My" `
            -Signer $rootCA -NotAfter (Get-Date).AddYears($CertValidityYears)
        $blazorCertThumbprint = $blazorCert.Thumbprint
        Write-OK "Issued Blazor cert: $blazorCertThumbprint"

        $blazorPfxPath = Join-Path $CertDir 'blazor.pfx'
        $pfxPwd = ConvertTo-SecureString -String $blazorPfxPassword -Force -AsPlainText
        Export-PfxCertificate -Cert $blazorCert -FilePath $blazorPfxPath -Password $pfxPwd | Out-Null
        Write-OK "Exported Blazor PFX: $blazorPfxPath"
    }

    'Mkcert' {
        $mkcert = Get-Command mkcert -ErrorAction SilentlyContinue
        if (-not $mkcert) {
            throw "mkcert.exe not found on PATH. Install via scoop: `scoop install mkcert` and re-run."
        }

        # Install root CA (idempotent — mkcert -install no-ops if already present)
        & mkcert -install | Out-Null
        Write-OK "mkcert root CA installed in LocalMachine\Root"

        Push-Location $CertDir
        try {
            & mkcert -cert-file sqlserver.crt -key-file sqlserver.key $sqlSans | Out-Null
            # SQL Server on Windows reads from cert store, not files — convert PEM to PFX and import
            & openssl pkcs12 -export `
                -out sqlserver.pfx `
                -inkey sqlserver.key -in sqlserver.crt `
                -password "pass:changeit" | Out-Null
            $sqlImport = Import-PfxCertificate `
                -FilePath (Join-Path $CertDir 'sqlserver.pfx') `
                -CertStoreLocation 'Cert:\LocalMachine\My' `
                -Password (ConvertTo-SecureString 'changeit' -AsPlainText -Force) `
                -Exportable
            $sqlCertThumbprint = $sqlImport.Thumbprint
            Write-OK "Imported SQL cert: $sqlCertThumbprint"

            & mkcert -pkcs12 -p12-file blazor.pfx $blazorSans | Out-Null
            $blazorPfxPath = Join-Path $CertDir 'blazor.pfx'
            $blazorImport = Import-PfxCertificate `
                -FilePath $blazorPfxPath `
                -CertStoreLocation 'Cert:\LocalMachine\My' `
                -Password (ConvertTo-SecureString $blazorPfxPassword -AsPlainText -Force) `
                -Exportable
            $blazorCertThumbprint = $blazorImport.Thumbprint
            Write-OK "Imported Blazor cert: $blazorCertThumbprint (PFX: $blazorPfxPath)"
        }
        finally { Pop-Location }
    }

    'Existing' {
        $sqlCertThumbprint = $ExistingSqlCertThumbprint
        $blazorCertThumbprint = $ExistingBlazorCertThumbprint

        $sqlCert = Get-ChildItem "Cert:\LocalMachine\My\$sqlCertThumbprint" -ErrorAction SilentlyContinue
        $blazorCert = Get-ChildItem "Cert:\LocalMachine\My\$blazorCertThumbprint" -ErrorAction SilentlyContinue
        if (-not $sqlCert) { throw "SQL cert thumbprint $sqlCertThumbprint not in Cert:\LocalMachine\My" }
        if (-not $blazorCert) { throw "Blazor cert thumbprint $blazorCertThumbprint not in Cert:\LocalMachine\My" }

        # Export Blazor cert as PFX in case Kestrel needs it
        $blazorPfxPath = Join-Path $CertDir 'blazor.pfx'
        $pfxPwd = ConvertTo-SecureString -String $blazorPfxPassword -Force -AsPlainText
        Export-PfxCertificate -Cert $blazorCert -FilePath $blazorPfxPath -Password $pfxPwd | Out-Null
        Write-OK "Re-exported existing Blazor cert to $blazorPfxPath (for Kestrel hosting mode)"
    }
}

# ──────────────────────────────────────────────────────────────────
# 4. Wire SQL Server (cert ACL + registry + restart)
# ──────────────────────────────────────────────────────────────────
if ($sqlSvc) {
    Write-Step "4. Wiring SQL Server to use the new cert"

    # 4a. Grant SQL Server's service account READ on the cert private key
    $svcAccount = (Get-WmiObject Win32_Service -Filter "Name='$($sqlSvc[0].Name)'").StartName
    Write-OK "SQL Server runs as: $svcAccount"

    $cert = Get-ChildItem "Cert:\LocalMachine\My\$sqlCertThumbprint"
    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
    if ($rsa.Key.UniqueName) {
        $keyFile = $rsa.Key.UniqueName
        $candidates = @(
            "$env:ProgramData\Microsoft\Crypto\Keys\$keyFile",
            "$env:ProgramData\Microsoft\Crypto\RSA\MachineKeys\$keyFile"
        )
        $keyPath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
        if ($keyPath) {
            & icacls $keyPath /grant "${svcAccount}:R" | Out-Null
            Write-OK "Granted READ on $keyPath to $svcAccount"
        } else {
            Write-Warn2 "Could not locate private-key file for thumbprint $sqlCertThumbprint — SQL Server may fail to start. Check `$env:ProgramData\Microsoft\Crypto\` manually."
        }
    }

    # 4b. Set Certificate + ForceEncryption in SQL Server registry
    $instanceId = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL")."MSSQLSERVER"
    if (-not $instanceId) {
        # Find the first instance under "Instance Names\SQL"
        $instanceProps = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL"
        $instanceId = $instanceProps.PSObject.Properties |
            Where-Object { $_.Name -notlike 'PS*' } |
            Select-Object -First 1 -ExpandProperty Value
    }
    $instanceKey = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$instanceId\MSSQLServer\SuperSocketNetLib"
    if (Test-Path $instanceKey) {
        Set-ItemProperty -Path $instanceKey -Name "Certificate" -Value $sqlCertThumbprint.ToLower()
        Set-ItemProperty -Path $instanceKey -Name "ForceEncryption" -Value 1
        Write-OK "Wrote cert thumbprint + ForceEncryption=1 to $instanceKey"

        # 4c. Restart SQL Server
        if ($PSCmdlet.ShouldProcess($sqlSvc[0].Name, "Restart-Service")) {
            Restart-Service -Name $sqlSvc[0].Name -Force
            Write-OK "Restarted $($sqlSvc[0].Name)"
        }
    } else {
        Write-Warn2 "SQL Server registry path not found: $instanceKey. Skipping cert wiring; configure manually via SQL Server Configuration Manager."
    }
} else {
    Write-Step "4. Skipping SQL Server cert wiring (no local SQL Server service)"
    Write-Warn2 "If SQL Server is on a different host, run §2 of WindowsServerCertSetup.md there."
}

# ──────────────────────────────────────────────────────────────────
# 5. Wire Blazor (Kestrel env vars OR IIS site + binding)
# ──────────────────────────────────────────────────────────────────
Write-Step "5. Wiring Blazor host ($HostingModel)"

if ($HostingModel -eq 'Kestrel') {
    # Set machine-scope env vars so the app pool / scheduled task /
    # service launcher picks them up regardless of how it's started.
    [Environment]::SetEnvironmentVariable("ASPNETCORE_Kestrel__Certificates__Default__Path", $blazorPfxPath, "Machine")
    [Environment]::SetEnvironmentVariable("ASPNETCORE_Kestrel__Certificates__Default__Password", $blazorPfxPassword, "Machine")
    [Environment]::SetEnvironmentVariable("ASPNETCORE_URLS", "https://*:$BlazorHttpsPort;http://*:$BlazorHttpPort", "Machine")
    [Environment]::SetEnvironmentVariable("ASPNETCORE_ENVIRONMENT", "Production", "Machine")
    Write-OK "Set ASPNETCORE_* machine env vars"
    Write-OK "Start the app via: cd $appDir\bin\Debug\net10.0; dotnet HomeFrontPOC.dll"
}
else {
    Import-Module WebAdministration
    $siteName = 'HomeFrontPOC'

    # 5a. Site (idempotent — remove + recreate)
    if (Get-Website -Name $siteName -ErrorAction SilentlyContinue) {
        Remove-Website -Name $siteName
        Write-OK "Removed existing IIS site $siteName"
    }
    New-Website -Name $siteName `
        -PhysicalPath (Join-Path $appDir 'bin\Debug\net10.0') `
        -Port 80 -HostHeader $SqlServer | Out-Null
    Write-OK "Created IIS site $siteName"

    # 5b. HTTPS binding with the cert
    New-WebBinding -Name $siteName -Protocol https -Port 443 `
        -HostHeader $SqlServer -SslFlags 1 | Out-Null
    $binding = Get-WebBinding -Name $siteName -Protocol https
    $binding.AddSslCertificate($blazorCertThumbprint, "My") | Out-Null
    Write-OK "Added HTTPS binding (cert $blazorCertThumbprint)"

    # 5c. App pool: no managed code (Kestrel handles .NET, IIS just proxies)
    Set-ItemProperty "IIS:\AppPools\$siteName" -Name managedRuntimeVersion -Value ""
    if ($AppServiceAccount) {
        # Set app pool identity to a domain account
        Set-ItemProperty "IIS:\AppPools\$siteName" -Name processModel `
            -Value @{userName=$AppServiceAccount; identityType="SpecificUser"}
        Write-OK "App pool identity set to $AppServiceAccount"
    }

    # 5d. In-process hosting model — write web.config snippet
    if ($HostingModel -eq 'IisInProcess') {
        $webConfig = Join-Path $appDir 'bin\Debug\net10.0\web.config'
        if (Test-Path $webConfig) {
            $xml = [xml](Get-Content $webConfig)
            $node = $xml.configuration.'system.webServer'.aspNetCore
            if ($node) {
                $node.SetAttribute('hostingModel', 'InProcess')
                $xml.Save($webConfig)
                Write-OK "Set hostingModel=InProcess in web.config"
            }
        }
    }

    Write-OK "IIS site $siteName ready on https://$SqlServer/"
}

# ──────────────────────────────────────────────────────────────────
# 6. Configure connection string + password source
# ──────────────────────────────────────────────────────────────────
Write-Step "6. Configuring connection string + password source"

$appsettingsPath = Join-Path $appDir 'bin\Debug\net10.0\appsettings.json'
if (-not (Test-Path $appsettingsPath)) {
    throw "appsettings.json not found at $appsettingsPath"
}
$settings = Get-Content $appsettingsPath -Raw | ConvertFrom-Json

if ($UseIntegratedSecurity) {
    $connStr = "Server=$SqlServer,$SqlPort;Database=$DatabaseName;Integrated Security=True;Encrypt=False;TrustServerCertificate=true;"
    Write-OK "Connection string: Integrated Security (no password)"
    Write-Warn2 "DBA must run on the SQL Server:"
    Write-Host "  CREATE LOGIN [$AppServiceAccount] FROM WINDOWS;" -ForegroundColor Gray
    Write-Host "  USE [$DatabaseName];" -ForegroundColor Gray
    Write-Host "  CREATE USER [$AppServiceAccount] FOR LOGIN [$AppServiceAccount];" -ForegroundColor Gray
    Write-Host "  EXEC sp_addrolemember 'db_owner', '$AppServiceAccount';" -ForegroundColor Gray
} else {
    $connStr = "Server=$SqlServer,$SqlPort;Database=$DatabaseName;User Id=sa;TrustServerCertificate=true;"
    [Environment]::SetEnvironmentVariable("Database__Password", $SaPassword, "Machine")
    Write-OK "Connection string: SQL auth (User Id=sa, password via env var)"
    Write-OK "Set machine env var Database__Password (double underscore)"
}

$settings.ConnectionStrings.DefaultConnection = $connStr
$settings | ConvertTo-Json -Depth 32 | Set-Content $appsettingsPath -Encoding UTF8
Write-OK "Updated $appsettingsPath"

# Restrict appsettings ACL
& icacls $appsettingsPath /inheritance:r 2>&1 | Out-Null
& icacls $appsettingsPath /grant "${AppServiceAccount}:R" "BUILTIN\Administrators:F" 2>&1 | Out-Null
Write-OK "Restricted appsettings.json ACL to $AppServiceAccount + Admins"

# ──────────────────────────────────────────────────────────────────
# 7. Smoke test
# ──────────────────────────────────────────────────────────────────
if ($SkipSmokeTest) {
    Write-Step "7. Skipping smoke test (-SkipSmokeTest)"
} else {
    Write-Step "7. Smoke test"

    $sqlcmd = Get-Command sqlcmd -ErrorAction SilentlyContinue
    if (-not $sqlcmd) {
        Write-Warn2 "sqlcmd not on PATH — skipping SQL TLS check. Install with: choco install sqlserver-cmdlineutils"
    } else {
        try {
            if ($UseIntegratedSecurity) {
                $output = & sqlcmd -S "$SqlServer,$SqlPort" -E -N -Q "SELECT 'OK' AS Verdict, encrypt_option FROM sys.dm_exec_connections WHERE session_id = @@SPID" 2>&1
            } else {
                $output = & sqlcmd -S "$SqlServer,$SqlPort" -U sa -P $SaPassword -N -Q "SELECT 'OK' AS Verdict, encrypt_option FROM sys.dm_exec_connections WHERE session_id = @@SPID" 2>&1
            }
            if ($LASTEXITCODE -eq 0) {
                Write-OK "SQL strict TLS test: PASSED"
                $output | ForEach-Object { Write-Host "      $_" -ForegroundColor Gray }
            } else {
                Write-Warn2 "SQL strict TLS test: FAILED ($LASTEXITCODE)"
                $output | ForEach-Object { Write-Host "      $_" -ForegroundColor Red }
            }
        } catch {
            Write-Warn2 "SQL strict TLS test threw: $_"
        }
    }

    if ($HostingModel -eq 'Kestrel') {
        Write-Warn2 "Blazor HTTPS test skipped — start the app first then run:"
        Write-Host "      Invoke-WebRequest -Uri https://${SqlServer}:$BlazorHttpsPort -UseBasicParsing" -ForegroundColor Gray
    } else {
        try {
            $response = Invoke-WebRequest -Uri "https://$SqlServer/" -UseBasicParsing -ErrorAction Stop
            if ($response.StatusCode -eq 200) {
                Write-OK "Blazor HTTPS test: 200 OK on https://$SqlServer/"
            } else {
                Write-Warn2 "Blazor HTTPS returned $($response.StatusCode)"
            }
        } catch {
            Write-Warn2 "Blazor HTTPS test failed: $_"
        }
    }
}

# ──────────────────────────────────────────────────────────────────
# 8. Summary
# ──────────────────────────────────────────────────────────────────
Write-Step "Done"
Write-Host ""
Write-Host "  App installed:   $appDir" -ForegroundColor White
Write-Host "  Cert files:      $CertDir" -ForegroundColor White
if ($sqlCertThumbprint) {
    Write-Host "  SQL cert:        $sqlCertThumbprint" -ForegroundColor White
}
if ($blazorCertThumbprint) {
    Write-Host "  Blazor cert:     $blazorCertThumbprint" -ForegroundColor White
}
Write-Host "  SQL Server:      $SqlServer,$SqlPort  Database=$DatabaseName" -ForegroundColor White
if ($UseIntegratedSecurity) {
    Write-Host "  Auth mode:       Integrated Security (Windows)" -ForegroundColor White
    Write-Host "  App account:     $AppServiceAccount" -ForegroundColor White
} else {
    Write-Host "  Auth mode:       SQL Server (sa + env var Database__Password)" -ForegroundColor White
}
Write-Host "  Hosting:         $HostingModel" -ForegroundColor White
Write-Host ""

if ($HostingModel -eq 'Kestrel') {
    Write-Host "Next: start the app from a NEW shell so env vars take effect:" -ForegroundColor Cyan
    Write-Host "  cd $appDir\bin\Debug\net10.0" -ForegroundColor Gray
    Write-Host "  dotnet HomeFrontPOC.dll" -ForegroundColor Gray
} else {
    Write-Host "Next: browse to https://$SqlServer/login to verify." -ForegroundColor Cyan
}
Write-Host ""
Write-Host "Welcome-screen toggles (after first login):" -ForegroundColor Cyan
Write-Host "  • CHECK 'Encrypt DB connection (TLS)' — proves the SQL cert chain"
if ($UseIntegratedSecurity) {
    Write-Host "  • 'Get DB Password from System Secrets' — no-op under Integrated Security"
} else {
    Write-Host "  • 'Get DB Password from System Secrets' — leave CHECKED (env var path)"
}
Write-Host ""
