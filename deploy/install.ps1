#Requires -Version 5.1
<#
.SYNOPSIS
    Instal·la RichCell a un nou ordinador.
.DESCRIPTION
    1. Copia els fitxers del projecte a C:\RichCell
    2. Instal·la Node.js si cal (obre el navegador)
    3. Instal·la paquets npm
    4. Instal·la el certificat CA al magatzem de confiança
    5. Crea una tasca programada per arrencar el servidor automàticament
    6. Comparteix la carpeta local per al catàleg d'Excel
    7. Configura Excel per mostrar el complement
.EXAMPLE
    # Des de la carpeta compartida de xarxa:
    powershell -ExecutionPolicy Bypass -File "\\NOM-PC\RichCell\deploy\install.ps1"
#>

$ErrorActionPreference = "Stop"
$DestPath = "C:\RichCell"
$TaskName = "RichCell_Server"

# ── Colors ────────────────────────────────────────────────────────────────────
function OK   { param($msg) Write-Host "  [OK] $msg" -ForegroundColor Green }
function INFO { param($msg) Write-Host "  [..] $msg" -ForegroundColor Cyan }
function WARN { param($msg) Write-Host "  [!!] $msg" -ForegroundColor Yellow }
function FAIL { param($msg) Write-Host "  [XX] $msg" -ForegroundColor Red }

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║   RichCell  –  Instal·lació                      ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# ── 1. Copia fitxers ──────────────────────────────────────────────────────────
INFO "Copiant fitxers del projecte..."
$SourceRoot = Split-Path -Parent $PSScriptRoot   # un nivell amunt de /deploy

if (-not (Test-Path $DestPath)) {
    New-Item -ItemType Directory -Path $DestPath | Out-Null
}

# Copia tot excepte la carpeta deploy (evita recursió)
Get-ChildItem -Path $SourceRoot -Exclude "deploy","node_modules" | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination $DestPath -Recurse -Force
}
OK "Fitxers copiats a $DestPath"

# ── 2. Comprova Node.js ───────────────────────────────────────────────────────
INFO "Comprovant Node.js..."
$nodePath = "C:\Program Files\nodejs\node.exe"
if (-not (Test-Path $nodePath)) {
    FAIL "Node.js no trobat."
    Write-Host ""
    Write-Host "  Instal·la Node.js LTS des de: https://nodejs.org" -ForegroundColor Yellow
    Write-Host "  Un cop instal·lat, torna a executar aquest script." -ForegroundColor Yellow
    Start-Process "https://nodejs.org/en/download"
    exit 1
}
$nodeVer = & $nodePath --version 2>&1
OK "Node.js $nodeVer"

# ── 3. Instal·la paquets npm ──────────────────────────────────────────────────
INFO "Instal·lant paquets npm..."
$npm = "C:\Program Files\nodejs\npm.cmd"
Push-Location $DestPath
& $npm install --quiet 2>&1 | Out-Null
Pop-Location
OK "Paquets npm instal·lats"

# ── 4. Instal·la certificat CA ────────────────────────────────────────────────
INFO "Instal·lant certificat CA..."
$caCert = Join-Path $DestPath "certs\ca.crt"

if (-not (Test-Path $caCert)) {
    FAIL "No es troba $caCert — assegura't que els certs s'han copiat correctament."
    exit 1
}

# Comprova si ja és al magatzem
$thumbprint = (Get-PfxCertificate -FilePath $caCert -ErrorAction SilentlyContinue).Thumbprint
$alreadyLM  = Get-ChildItem Cert:\LocalMachine\Root -ErrorAction SilentlyContinue |
              Where-Object { $_.Thumbprint -eq $thumbprint }

if ($alreadyLM) {
    OK "CA ja instal·lat a LocalMachine\Root"
} else {
    # Intenta instal·lar (cal admin — auto-elevació)
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
                [Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($isAdmin) {
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root","LocalMachine")
        $store.Open("ReadWrite")
        $store.Add((New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($caCert)))
        $store.Close()
        OK "CA instal·lat a LocalMachine\Root"
    } else {
        WARN "Necessita privilegis d'administrador per instal·lar el CA."
        INFO "Llançant finestra elevada..."
        $args = "-ExecutionPolicy Bypass -Command `"
            `$store = New-Object System.Security.Cryptography.X509Certificates.X509Store('Root','LocalMachine');
            `$store.Open('ReadWrite');
            `$store.Add((New-Object System.Security.Cryptography.X509Certificates.X509Certificate2('$caCert')));
            `$store.Close();
            Write-Host 'CA instal·lat correctament.' -ForegroundColor Green;
            Start-Sleep 2
        `""
        Start-Process powershell -ArgumentList $args -Verb RunAs -Wait
        OK "CA instal·lat a LocalMachine\Root (finestra elevada)"
    }
}

# Instal·la també a CurrentUser (sense admin)
$storeCU = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root","CurrentUser")
$storeCU.Open("ReadWrite")
$storeCU.Add((New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($caCert)))
$storeCU.Close()
OK "CA instal·lat a CurrentUser\Root"

# ── 5. Tasca programada (servidor arrenca amb Windows) ────────────────────────
INFO "Configurant tasca programada per al servidor..."
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

$action   = New-ScheduledTaskAction `
                -Execute  $nodePath `
                -Argument "server.js" `
                -WorkingDirectory $DestPath

$trigger  = New-ScheduledTaskTrigger -AtLogon

$settings = New-ScheduledTaskSettingsSet `
                -ExecutionTimeLimit (New-TimeSpan -Hours 0) `
                -RestartCount 3 `
                -RestartInterval (New-TimeSpan -Minutes 1) `
                -StartWhenAvailable

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action   $action `
    -Trigger  $trigger `
    -Settings $settings `
    -RunLevel Highest `
    -Force | Out-Null

# Arrenca ara mateix
Start-ScheduledTask -TaskName $TaskName
Start-Sleep -Seconds 3
OK "Servidor arrencat (s'iniciarà automàticament en cada inici de sessió)"

# ── 6. Comparteix la carpeta local ────────────────────────────────────────────
INFO "Compartint carpeta local..."
$shareName = "RichCell"
$shareExists = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
if (-not $shareExists) {
    $everyone = (New-Object System.Security.Principal.SecurityIdentifier("S-1-1-0")).Translate(
                    [System.Security.Principal.NTAccount]).Value
    New-SmbShare -Name $shareName -Path $DestPath -FullAccess $everyone | Out-Null
    OK "Carpeta compartida: \\$env:COMPUTERNAME\$shareName"
} else {
    OK "Compartició ja existent: \\$env:COMPUTERNAME\$shareName"
}
$uncPath = "\\$env:COMPUTERNAME\$shareName"

# ── 7. Configura catàleg d'Excel via registre ─────────────────────────────────
INFO "Configurant catàleg de confiança a Excel..."

# Neteja entrades anteriors amb la mateixa URL
$catalogsKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
if (Test-Path $catalogsKey) {
    Get-ChildItem $catalogsKey | ForEach-Object {
        $url = (Get-ItemProperty $_.PSPath -Name Url -ErrorAction SilentlyContinue).Url
        if ($url -and $url -like "*RichCell*") {
            Remove-Item $_.PSPath -Force
        }
    }
}

# Crea nova entrada
$guid    = [System.Guid]::NewGuid().ToString("B").ToUpper()
$regPath = "$catalogsKey\$guid"
New-Item  -Path $regPath -Force | Out-Null
Set-ItemProperty -Path $regPath -Name "Url"   -Value $uncPath
Set-ItemProperty -Path $regPath -Name "Flags" -Value 1 -Type DWord
OK "Catàleg configurat al registre"

# ── Resum final ───────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "╔══════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║   Instal·lació completada correctament!          ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Write-Host "  Passos finals:" -ForegroundColor White
Write-Host "  1. Tanca Excel completament (si estava obert)"   -ForegroundColor Yellow
Write-Host "  2. Torna a obrir Excel"                          -ForegroundColor Yellow
Write-Host "  3. Inserir → Complements → CARPETA COMPARTIDA"   -ForegroundColor Yellow
Write-Host "  4. Selecciona 'RichCell'"                        -ForegroundColor Yellow
Write-Host ""
