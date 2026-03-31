Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =========================
# CONFIG MAIL GMAIL
# =========================
$smtpServer = "smtp.gmail.com"
$smtpPort = 587
$smtpSsl = $true

$mailFrom = "gerald.laronche@gmail.com"
$mailTo = "jllsla@free.fr"
$mailUser = "gerald.laronche@gmail.com"
$mailPass = "yuttmuagkfqhyaxd"

# =========================
# CONFIG APP / UPDATE
# =========================
$AppVersion = "1.2"
$GitHubLatestApiUrl = "https://api.github.com/repos/Pinkywhisky/diagnostic-pc/releases/latest"
$SetupAssetName = "Setup_Diagnostic_PC.exe"

# =========================
# VARIABLES GLOBALES
# =========================
$script:LastReport = ""
$script:UpdateChecked = $false

# =========================
# FENETRE PRINCIPALE
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Diagnostic PC v$AppVersion"
$form.Size = New-Object System.Drawing.Size(1000, 720)
$form.MinimumSize = New-Object System.Drawing.Size(760, 520)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White

# =========================
# HEADER
# =========================
$panelTop = New-Object System.Windows.Forms.Panel
$panelTop.Dock = "Top"
$panelTop.Height = 95
$panelTop.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

$labelTitle = New-Object System.Windows.Forms.Label
$labelTitle.Text = "Diagnostic PC"
$labelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$labelTitle.AutoSize = $true
$labelTitle.Location = New-Object System.Drawing.Point(16, 12)

$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Text = "Version $AppVersion"
$labelVersion.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$labelVersion.ForeColor = [System.Drawing.Color]::DimGray
$labelVersion.AutoSize = $true
$labelVersion.Location = New-Object System.Drawing.Point(19, 48)

$labelInfo = New-Object System.Windows.Forms.Label
$labelInfo.Text = "1. Lancez le diagnostic   2. Envoyez le résultat par mail"
$labelInfo.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$labelInfo.AutoSize = $true
$labelInfo.Location = New-Object System.Drawing.Point(16, 68)

$panelTop.Controls.Add($labelTitle)
$panelTop.Controls.Add($labelVersion)
$panelTop.Controls.Add($labelInfo)

# =========================
# ZONE TEXTE
# =========================
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.ScrollBars = "Vertical"
$textBox.ReadOnly = $true
$textBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$textBox.Dock = "Fill"
$textBox.BackColor = [System.Drawing.Color]::White
$textBox.BorderStyle = "FixedSingle"

$panelMain = New-Object System.Windows.Forms.Panel
$panelMain.Dock = "Fill"
$panelMain.Padding = New-Object System.Windows.Forms.Padding(12)
$panelMain.Controls.Add($textBox)

# =========================
# PANEL BAS
# =========================
$panelBottom = New-Object System.Windows.Forms.Panel
$panelBottom.Height = 70
$panelBottom.Dock = "Bottom"
$panelBottom.BackColor = [System.Drawing.Color]::FromArgb(248, 248, 248)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Prêt"
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$statusLabel.AutoSize = $true
$statusLabel.Location = New-Object System.Drawing.Point(16, 12)
$statusLabel.ForeColor = [System.Drawing.Color]::DimGray

# =========================
# BOUTONS
# =========================
$buttonRun = New-Object System.Windows.Forms.Button
$buttonRun.Text = "Lancer le diagnostic"
$buttonRun.Size = New-Object System.Drawing.Size(180, 34)
$buttonRun.Location = New-Object System.Drawing.Point(16, 28)
$buttonRun.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$buttonRun.FlatStyle = "Flat"
$buttonRun.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$buttonRun.ForeColor = [System.Drawing.Color]::White
$buttonRun.FlatAppearance.BorderSize = 0

$buttonMail = New-Object System.Windows.Forms.Button
$buttonMail.Text = "Envoyer par mail"
$buttonMail.Size = New-Object System.Drawing.Size(140, 34)
$buttonMail.Location = New-Object System.Drawing.Point(210, 28)
$buttonMail.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$buttonMail.FlatStyle = "Flat"
$buttonMail.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$buttonMail.FlatAppearance.BorderColor = [System.Drawing.Color]::Silver

$buttonClose = New-Object System.Windows.Forms.Button
$buttonClose.Text = "Fermer"
$buttonClose.Size = New-Object System.Drawing.Size(100, 34)
$buttonClose.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$buttonClose.FlatStyle = "Flat"
$buttonClose.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$buttonClose.FlatAppearance.BorderColor = [System.Drawing.Color]::Silver

$panelBottom.Add_Resize({
        $buttonClose.Location = New-Object System.Drawing.Point(($panelBottom.ClientSize.Width - $buttonClose.Width - 16), 28)
    })

$buttonClose.Location = New-Object System.Drawing.Point(($panelBottom.ClientSize.Width - $buttonClose.Width - 16), 28)

$panelBottom.Controls.Add($statusLabel)
$panelBottom.Controls.Add($buttonRun)
$panelBottom.Controls.Add($buttonMail)
$panelBottom.Controls.Add($buttonClose)

$form.Controls.Add($panelMain)
$form.Controls.Add($panelBottom)
$form.Controls.Add($panelTop)

$form.AcceptButton = $buttonRun
$form.CancelButton = $buttonClose

# =========================
# FONCTIONS
# =========================
function Set-Status {
    param(
        [string]$Text
    )

    $statusLabel.Text = $Text
    $statusLabel.Refresh()
}

function Add-TextLine {
    param([string]$Text)
    $textBox.AppendText($Text + [Environment]::NewLine)
    $textBox.SelectionStart = $textBox.Text.Length
    $textBox.ScrollToCaret()
}

function Add-Separator {
    Add-TextLine "=============================="
}

function Get-UptimeString {
    param([datetime]$LastBoot)

    $uptime = (Get-Date) - $LastBoot
    return "{0} jour(s) {1} heure(s) {2} minute(s)" -f $uptime.Days, $uptime.Hours, $uptime.Minutes
}

function Add-CriticalProblem {
    param(
        [ref]$ProblemList,
        [string]$Message
    )

    if (-not [string]::IsNullOrWhiteSpace($Message) -and -not $ProblemList.Value.Contains($Message)) {
        $ProblemList.Value.Add($Message) | Out-Null
    }
}

function Add-WarningProblem {
    param(
        [ref]$ProblemList,
        [string]$Message
    )

    if (-not [string]::IsNullOrWhiteSpace($Message) -and -not $ProblemList.Value.Contains($Message)) {
        $ProblemList.Value.Add($Message) | Out-Null
    }
}

function Get-NormalizedVersionString {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VersionText
    )

    $normalized = $VersionText.Trim()

    if ($normalized.StartsWith("v", [System.StringComparison]::OrdinalIgnoreCase)) {
        $normalized = $normalized.Substring(1)
    }

    return $normalized
}

function Get-LatestReleaseInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ApiUrl,
        [Parameter(Mandatory = $true)]
        [string]$CurrentVersion,
        [Parameter(Mandatory = $true)]
        [string]$ExpectedAssetName
    )

    try {
        $headers = @{
            "User-Agent" = "Diagnostic-PC"
        }

        $release = Invoke-RestMethod -Uri $ApiUrl -Headers $headers -Method Get -ErrorAction Stop

        $latestTag = $release.tag_name
        if ([string]::IsNullOrWhiteSpace($latestTag)) {
            throw "Impossible de lire le tag de la dernière release."
        }

        $currentNormalized = Get-NormalizedVersionString -VersionText $CurrentVersion
        $latestNormalized = Get-NormalizedVersionString -VersionText $latestTag

        $currentVersionObj = [version]$currentNormalized
        $latestVersionObj = [version]$latestNormalized

        $asset = $release.assets | Where-Object { $_.name -eq $ExpectedAssetName } | Select-Object -First 1
        if (-not $asset) {
            throw "Impossible de trouver l'asset '$ExpectedAssetName' dans la dernière release."
        }

        return [pscustomobject]@{
            Success         = $true
            CurrentVersion  = $currentNormalized
            LatestVersion   = $latestNormalized
            LatestTag       = $latestTag
            UpdateAvailable = ($latestVersionObj -gt $currentVersionObj)
            AssetName       = $asset.name
            AssetUrl        = $asset.browser_download_url
        }
    }
    catch {
        return [pscustomobject]@{
            Success         = $false
            ErrorMessage    = $_.Exception.Message
            UpdateAvailable = $false
        }
    }
}

function Start-AppUpdate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DownloadUrl,
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    if ([string]::IsNullOrWhiteSpace($DownloadUrl)) {
        throw "URL de téléchargement vide."
    }

    if ([string]::IsNullOrWhiteSpace($FileName)) {
        throw "Nom de fichier vide."
    }

    if (-not $FileName.EndsWith(".exe", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Le fichier de mise à jour n'est pas un exécutable valide."
    }

    $tempDir = Join-Path $env:TEMP "DiagnosticPC"
    if (-not (Test-Path $tempDir)) {
        New-Item -ItemType Directory -Path $tempDir | Out-Null
    }

    $tempFile = Join-Path $tempDir $FileName

    if (Test-Path $tempFile) {
        Remove-Item $tempFile -Force
    }

    Invoke-WebRequest -Uri $DownloadUrl -OutFile $tempFile -UseBasicParsing -ErrorAction Stop

    if (-not (Test-Path $tempFile)) {
        throw "Le fichier téléchargé est introuvable : $tempFile"
    }

    $fileInfo = Get-Item $tempFile
    if ($fileInfo.Length -le 0) {
        throw "Le fichier téléchargé est vide."
    }

    return $tempFile
}

function Send-DiagMail {
    param(
        [string]$Subject,
        [string]$Body
    )

    try {
        if ([string]::IsNullOrWhiteSpace($mailPass)) {
            throw "Mot de passe Gmail non configuré."
        }

        $message = New-Object System.Net.Mail.MailMessage
        $message.From = $mailFrom
        $message.To.Add($mailTo)
        $message.Subject = $Subject
        $message.Body = $Body
        $message.IsBodyHtml = $false
        $message.BodyEncoding = [System.Text.Encoding]::UTF8
        $message.SubjectEncoding = [System.Text.Encoding]::UTF8

        $smtp = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
        $smtp.EnableSsl = $smtpSsl
        $smtp.Credentials = New-Object System.Net.NetworkCredential($mailUser, $mailPass)

        $smtp.Send($message)

        [System.Windows.Forms.MessageBox]::Show(
            "Mail envoyé avec succès.",
            "Envoi réussi",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Erreur lors de l'envoi du mail : $($_.Exception.Message)",
            "Erreur d'envoi",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

function Test-TcpPort {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HostName,

        [Parameter(Mandatory = $true)]
        [int]$Port,

        [int]$TimeoutMs = 3000
    )

    $client = $null

    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $asyncResult = $client.BeginConnect($HostName, $Port, $null, $null)

        if (-not $asyncResult.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            return $false
        }

        $client.EndConnect($asyncResult)
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($client) {
            $client.Close()
            $client.Dispose()
        }
    }
}

function Start-Diag {
    $textBox.Clear()
    Set-Status "Diagnostic en cours..."

    $passerelle = $null
    $internet = "8.8.8.8"
    $dnsName = "google.com"
    $portTarget = "google.com"
    $ports = @(80, 443)

    $CriticalCount = 0
    $WarningCount = 0
    $CriticalProblems = [System.Collections.Generic.List[string]]::new()
    $WarningProblems = [System.Collections.Generic.List[string]]::new()

    Add-Separator
    Add-TextLine "       DIAGNOSTIC PC"
    Add-Separator
    Add-TextLine "Nom du poste : $env:COMPUTERNAME"
    Add-TextLine "Utilisateur  : $env:USERNAME"
    Add-TextLine "Date         : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')"
    Add-TextLine "Version outil : $AppVersion"
    Add-TextLine ""

    Add-TextLine "[SYSTEME]"
    try {
        $os = Get-CimInstance Win32_OperatingSystem
        Add-TextLine "OS           : $($os.Caption)"
        Add-TextLine "Version      : $($os.Version)"
        Add-TextLine "Build        : $($os.BuildNumber)"
        Add-TextLine "Dernier boot : $($os.LastBootUpTime.ToString('dd/MM/yyyy HH:mm:ss'))"
        Add-TextLine "Uptime       : $(Get-UptimeString -LastBoot $os.LastBootUpTime)"
    }
    catch {
        Add-TextLine "CRITICAL - Impossible de récupérer les informations système"
        $CriticalCount++
        Add-CriticalProblem ([ref]$CriticalProblems) "Impossible de récupérer les informations système"
    }
    Add-TextLine ""

    Add-TextLine "[BIOS]"
    try {
        $bios = Get-CimInstance Win32_BIOS
        $biosDate = if ($bios.ReleaseDate) { $bios.ReleaseDate.ToString('dd/MM/yyyy') } else { "Inconnue" }

        Add-TextLine "Fabricant : $($bios.Manufacturer)"
        Add-TextLine "Version   : $($bios.SMBIOSBIOSVersion)"
        Add-TextLine "Date      : $biosDate"
        Add-TextLine "Serial    : $($bios.SerialNumber)"
    }
    catch {
        Add-TextLine "CRITICAL - Impossible de récupérer les informations BIOS"
        $CriticalCount++
        Add-CriticalProblem ([ref]$CriticalProblems) "Impossible de récupérer les informations BIOS"
    }
    Add-TextLine ""

    Add-TextLine "[RESEAU LOCAL]"
    try {
        $netConfigs = Get-NetIPConfiguration | Where-Object {
            $_.IPv4Address -and $_.NetAdapter.Status -eq 'Up'
        }

        if ($netConfigs) {
            foreach ($cfg in $netConfigs) {
                $gatewayText = if ($cfg.IPv4DefaultGateway) { $cfg.IPv4DefaultGateway.NextHop } else { "Aucune" }

                Add-TextLine "Interface   : $($cfg.InterfaceAlias)"
                Add-TextLine "Description : $($cfg.NetAdapter.InterfaceDescription)"
                Add-TextLine "IPv4        : $($cfg.IPv4Address.IPAddress)"
                Add-TextLine "Passerelle  : $gatewayText"
                Add-TextLine "DNS         : $([string]::Join(', ', $cfg.DNSServer.ServerAddresses))"
                Add-TextLine ""

                if (-not $passerelle -and $cfg.IPv4DefaultGateway -and $cfg.IPv4DefaultGateway.NextHop) {
                    $passerelle = $cfg.IPv4DefaultGateway.NextHop
                }
            }
        }
        else {
            Add-TextLine "CRITICAL - Aucune interface IPv4 active trouvée"
            Add-TextLine ""
            $CriticalCount++
            Add-CriticalProblem ([ref]$CriticalProblems) "Aucune interface IPv4 active trouvée"
        }
    }
    catch {
        Add-TextLine "CRITICAL - Impossible de récupérer les informations réseau locales"
        Add-TextLine ""
        $CriticalCount++
        Add-CriticalProblem ([ref]$CriticalProblems) "Impossible de récupérer les informations réseau locales"
    }
    Add-TextLine ""
    
    Add-TextLine "[RESEAU]"

    if (Test-Connection -ComputerName $passerelle -Count 1 -Quiet -ErrorAction SilentlyContinue) {
        Add-TextLine "OK - Passerelle joignable ($passerelle)"
    }
    else {
        Add-TextLine "CRITICAL - Passerelle non joignable ($passerelle)"
        $CriticalCount++
        Add-CriticalProblem ([ref]$CriticalProblems) "Passerelle non joignable ($passerelle)"
    }

    if (Test-Connection -ComputerName $internet -Count 1 -Quiet -ErrorAction SilentlyContinue) {
        Add-TextLine "OK - Internet joignable ($internet)"
    }
    else {
        Add-TextLine "CRITICAL - Internet non joignable ($internet)"
        $CriticalCount++
        Add-CriticalProblem ([ref]$CriticalProblems) "Pas d'accès Internet ($internet non joignable)"
    }

    try {
        Resolve-DnsName -Name $dnsName -ErrorAction Stop | Out-Null
        Add-TextLine "OK - Résolution DNS fonctionnelle ($dnsName)"
    }
    catch {
        Add-TextLine "CRITICAL - Résolution DNS impossible ($dnsName)"
        $CriticalCount++
        Add-CriticalProblem ([ref]$CriticalProblems) "Résolution DNS impossible ($dnsName)"
    }
    Add-TextLine ""

    Add-TextLine "[PORTS]"
    foreach ($port in $ports) {
        try {
            $result = Test-TcpPort -HostName $portTarget -Port $port -TimeoutMs 3000

            if ($result) {
                Add-TextLine "OK - Port $port joignable sur $portTarget"
            }
            else {
                Add-TextLine "CRITICAL - Port $port non joignable sur $portTarget"
                $CriticalCount++
                Add-CriticalProblem ([ref]$CriticalProblems) "Port $port non joignable sur $portTarget"
            }
        }
        catch {
            Add-TextLine "CRITICAL - Erreur lors du test du port $port sur $portTarget"
            $CriticalCount++
            Add-CriticalProblem ([ref]$CriticalProblems) "Erreur lors du test du port $port sur $portTarget"
        }
    }
    Add-TextLine ""

    Add-TextLine "[CPU]"
    try {
        $cpu = Get-CimInstance Win32_Processor
        $cpuLoad = [math]::Round(($cpu | Measure-Object -Property LoadPercentage -Average).Average, 2)
        $cpuModel = ($cpu | Select-Object -First 1).Name
        $cpuCores = ($cpu | Measure-Object -Property NumberOfCores -Sum).Sum
        $cpuLogical = ($cpu | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum

        Add-TextLine "Modèle     : $cpuModel"
        Add-TextLine "Cœurs      : $cpuCores"
        Add-TextLine "Logiques   : $cpuLogical"
        Add-TextLine "Charge CPU : $cpuLoad %"

        if ($cpuLoad -ge 90) {
            Add-TextLine "CRITICAL - CPU très sollicité"
            $CriticalCount++
            Add-CriticalProblem ([ref]$CriticalProblems) "CPU très sollicité ($cpuLoad %)"
        }
    }
    catch {
        Add-TextLine "CRITICAL - Impossible de récupérer l'état CPU"
        $CriticalCount++
        Add-CriticalProblem ([ref]$CriticalProblems) "Impossible de récupérer l'état CPU"
    }
    Add-TextLine ""

    Add-TextLine "[RAM]"
    try {
        $osRam = Get-CimInstance Win32_OperatingSystem

        $totalRamGB = [math]::Round($osRam.TotalVisibleMemorySize / 1MB, 2)
        $freeRamGB = [math]::Round($osRam.FreePhysicalMemory / 1MB, 2)
        $usedRamGB = [math]::Round($totalRamGB - $freeRamGB, 2)

        if ($osRam.TotalVisibleMemorySize -gt 0) {
            $ramUsagePct = [math]::Round((($osRam.TotalVisibleMemorySize - $osRam.FreePhysicalMemory) / $osRam.TotalVisibleMemorySize) * 100, 2)
        }
        else {
            $ramUsagePct = 0
        }

        Add-TextLine "RAM totale   : $totalRamGB Go"
        Add-TextLine "RAM utilisée : $usedRamGB Go"
        Add-TextLine "RAM libre    : $freeRamGB Go"
        Add-TextLine "Utilisation  : $ramUsagePct %"

        if ($ramUsagePct -ge 90) {
            Add-TextLine "CRITICAL - RAM presque saturée"
            $CriticalCount++
            Add-CriticalProblem ([ref]$CriticalProblems) "RAM presque saturée ($ramUsagePct %)"
        }
    }
    catch {
        Add-TextLine "CRITICAL - Impossible de récupérer l'état mémoire"
        $CriticalCount++
        Add-CriticalProblem ([ref]$CriticalProblems) "Impossible de récupérer l'état mémoire"
    }
    Add-TextLine ""

    Add-TextLine "[DISQUES LOGIQUES]"
    try {
        $disques = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | Sort-Object DeviceID

        foreach ($d in $disques) {
            $tailleGo = [math]::Round($d.Size / 1GB, 2)
            $utiliseGo = [math]::Round(($d.Size - $d.FreeSpace) / 1GB, 2)
            $libreGo = [math]::Round($d.FreeSpace / 1GB, 2)

            if ($d.Size -gt 0) {
                $pct = [math]::Round((($d.Size - $d.FreeSpace) / $d.Size) * 100, 2)
            }
            else {
                $pct = 0
            }

            Add-TextLine "Lecteur $($d.DeviceID)"
            Add-TextLine "  Taille totale : $tailleGo Go"
            Add-TextLine "  Espace utilisé: $utiliseGo Go"
            Add-TextLine "  Espace libre  : $libreGo Go"
            Add-TextLine "  Utilisation   : $pct %"

            if ($pct -ge 90) {
                Add-TextLine "  CRITICAL      : disque presque plein"
                $CriticalCount++
                Add-CriticalProblem ([ref]$CriticalProblems) "Disque $($d.DeviceID) presque plein ($pct %)"
            }

            Add-TextLine ""
        }
    }
    catch {
        Add-TextLine "CRITICAL - Impossible de récupérer l'état des disques logiques"
        $CriticalCount++
        Add-CriticalProblem ([ref]$CriticalProblems) "Impossible de récupérer l'état des disques logiques"
        Add-TextLine ""
    }

    Add-TextLine "[DISQUES PHYSIQUES]"
    try {
        $physicalDisks = Get-CimInstance Win32_DiskDrive
        foreach ($disk in $physicalDisks) {
            $sizeGB = if ($disk.Size) { [math]::Round($disk.Size / 1GB, 2) } else { 0 }
            Add-TextLine "Modèle   : $($disk.Model)"
            Add-TextLine "Interface: $($disk.InterfaceType)"
            Add-TextLine "Taille   : $sizeGB Go"
            Add-TextLine "Statut   : $($disk.Status)"

            if ($disk.Status -and $disk.Status -ne "OK") {
                Add-TextLine "CRITICAL - Statut disque physique anormal"
                $CriticalCount++
                Add-CriticalProblem ([ref]$CriticalProblems) "Statut disque physique anormal pour $($disk.Model) : $($disk.Status)"
            }

            Add-TextLine ""
        }
    }
    catch {
        Add-TextLine "CRITICAL - Impossible de récupérer les informations des disques physiques"
        $CriticalCount++
        Add-CriticalProblem ([ref]$CriticalProblems) "Impossible de récupérer les informations des disques physiques"
        Add-TextLine ""
    }

    Add-TextLine "[SECURITE]"
    try {
        $defender = Get-Service -Name WinDefend -ErrorAction SilentlyContinue
        if ($defender -and $defender.Status -eq "Running") {
            Add-TextLine "OK - Microsoft Defender actif"
        }
        else {
            Add-TextLine "CRITICAL - Microsoft Defender inactif ou introuvable"
            $CriticalCount++
            Add-CriticalProblem ([ref]$CriticalProblems) "Microsoft Defender inactif ou introuvable"
        }
    }
    catch {
        Add-TextLine "CRITICAL - Impossible de vérifier Microsoft Defender"
        $CriticalCount++
        Add-CriticalProblem ([ref]$CriticalProblems) "Impossible de vérifier Microsoft Defender"
    }

    try {
        $fwProfiles = Get-NetFirewallProfile
        foreach ($fwProfile in $fwProfiles) {
            if ($fwProfile.Enabled) {
                Add-TextLine "OK - Firewall actif ($($fwProfile.Name))"
            }
            else {
                Add-TextLine "WARNING - Firewall désactivé ($($fwProfile.Name))"
                $WarningCount++
                Add-WarningProblem ([ref]$WarningProblems) "Firewall désactivé ($($fwProfile.Name))"
            }
        }
    }
    catch {
        Add-TextLine "WARNING - Impossible de vérifier le firewall"
        $WarningCount++
        Add-WarningProblem ([ref]$WarningProblems) "Impossible de vérifier le firewall"
    }
    Add-TextLine ""

    Add-TextLine "[SERVICES]"
    $services = @("Spooler", "wuauserv", "WinDefend")

    foreach ($serviceName in $services) {
        try {
            $srv = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
            if ($srv) {
                if ($srv.Status -eq "Running") {
                    Add-TextLine "OK - $serviceName actif"
                }
                else {
                    Add-TextLine "CRITICAL - $serviceName arrêté"
                    $CriticalCount++
                    Add-CriticalProblem ([ref]$CriticalProblems) "Service $serviceName arrêté"
                }
            }
            else {
                Add-TextLine "CRITICAL - $serviceName introuvable"
                $CriticalCount++
                Add-CriticalProblem ([ref]$CriticalProblems) "Service $serviceName introuvable"
            }
        }
        catch {
            Add-TextLine "CRITICAL - Impossible de vérifier $serviceName"
            $CriticalCount++
            Add-CriticalProblem ([ref]$CriticalProblems) "Impossible de vérifier le service $serviceName"
        }
    }
    Add-TextLine ""

    Add-Separator
    Add-TextLine "          RESULTAT"
    Add-Separator

    if ($CriticalCount -eq 0 -and $WarningCount -eq 0) {
        Add-TextLine "RESULTAT GLOBAL : OK"
    }
    elseif ($CriticalCount -eq 0 -and $WarningCount -gt 0) {
        Add-TextLine "RESULTAT GLOBAL : WARNING"
    }
    else {
        Add-TextLine "RESULTAT GLOBAL : CRITICAL"
    }

    Add-TextLine ""
    Add-TextLine "Résumé :"
    Add-TextLine "- Critical : $CriticalCount"
    Add-TextLine "- Warning  : $WarningCount"
    Add-TextLine ""

    if ($CriticalProblems.Count -gt 0) {
        Add-TextLine "PROBLEMES CRITIQUES :"
        Add-TextLine ""

        foreach ($probleme in $CriticalProblems) {
            Add-TextLine "- $probleme"
        }

        Add-TextLine ""
    }

    if ($WarningProblems.Count -gt 0) {
        Add-TextLine "AVERTISSEMENTS :"
        Add-TextLine ""

        foreach ($probleme in $WarningProblems) {
            Add-TextLine "- $probleme"
        }

        Add-TextLine ""
    }

    if ($CriticalProblems.Count -gt 0 -or $WarningProblems.Count -gt 0) {
        Add-TextLine "PISTES :"

        if ($CriticalProblems -contains "Passerelle non joignable ($passerelle)") {
            Add-TextLine "- Vérifier la connexion locale, l'IP du poste et la passerelle."
        }

        if ($CriticalProblems | Where-Object { $_ -like "Pas d'accès Internet*" }) {
            Add-TextLine "- Vérifier la connectivité WAN ou le chemin vers Internet."
        }

        if ($CriticalProblems | Where-Object { $_ -like "Résolution DNS impossible*" }) {
            Add-TextLine "- Vérifier les serveurs DNS configurés sur le poste."
        }

        if ($CriticalProblems | Where-Object { $_ -like "RAM presque saturée*" }) {
            Add-TextLine "- Fermer les applications gourmandes ou redémarrer le poste."
        }

        if ($CriticalProblems | Where-Object { $_ -like "CPU très sollicité*" }) {
            Add-TextLine "- Vérifier les processus fortement consommateurs de CPU."
        }

        if ($CriticalProblems | Where-Object { $_ -like "Disque * presque plein*" }) {
            Add-TextLine "- Libérer de l'espace disque sur le ou les lecteurs concernés."
        }

        if ($CriticalProblems | Where-Object { $_ -like "Port * non joignable*" }) {
            Add-TextLine "- Vérifier le firewall, le proxy ou la cible des tests de ports."
        }

        if ($CriticalProblems | Where-Object { $_ -like "Service * arrêté" }) {
            Add-TextLine "- Contrôler les services Windows signalés comme arrêtés."
        }

        if ($WarningProblems | Where-Object { $_ -like "Firewall désactivé*" }) {
            Add-TextLine "- Vérifier si la désactivation du firewall est volontaire."
        }
    }

    $script:LastReport = $textBox.Text
    Set-Status "Diagnostic terminé"
}

function Invoke-StartupUpdateCheck {
    try {
        Set-Status "Vérification des mises à jour..."
        $updateResult = Get-LatestReleaseInfo `
            -ApiUrl $GitHubLatestApiUrl `
            -CurrentVersion $AppVersion `
            -ExpectedAssetName $SetupAssetName

        if (-not $updateResult.Success) {
            Set-Status "Prêt"
            return
        }

        if (-not $updateResult.UpdateAvailable) {
            Set-Status "Prêt"
            return
        }

        $choice = [System.Windows.Forms.MessageBox]::Show(
            "Une nouvelle version est disponible : $($updateResult.LatestVersion)`r`nVersion actuelle : $($updateResult.CurrentVersion)`r`n`r`nVoulez-vous télécharger et lancer l'installation ?",
            "Mise à jour disponible",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
            Set-Status "Prêt"
            return
        }

        Set-Status "Téléchargement de la mise à jour..."
        $setupPath = Start-AppUpdate -DownloadUrl $updateResult.AssetUrl -FileName $updateResult.AssetName

        $finalChoice = [System.Windows.Forms.MessageBox]::Show(
            "La mise à jour a été téléchargée avec succès.`r`n`r`nVoulez-vous lancer l'installation maintenant ?",
            "Mise à jour prête",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        if ($finalChoice -ne [System.Windows.Forms.DialogResult]::Yes) {
            Set-Status "Prêt"
            return
        }

        Start-Process -FilePath $setupPath
        $form.Close()
    }
    catch {
        Set-Status "Prêt"
    }
}

# =========================
# ACTIONS
# =========================
$buttonRun.Add_Click({
        $buttonRun.Enabled = $false
        $buttonRun.Text = "En cours..."
        $form.Cursor = "WaitCursor"
        $form.Refresh()

        try {
            Start-Diag
        }
        finally {
            $buttonRun.Enabled = $true
            $buttonRun.Text = "Lancer le diagnostic"
            $form.Cursor = "Default"
        }
    })

$buttonMail.Add_Click({
        if ([string]::IsNullOrWhiteSpace($script:LastReport)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Aucun diagnostic à envoyer. Lance d'abord le diagnostic.",
                "Information",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        $confirmation = [System.Windows.Forms.MessageBox]::Show(
            "Autorisez-vous l'envoi du résultat à Gérald Laronche par mail ?",
            "Confirmation d'envoi",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
            $subject = "Diagnostic PC - $env:COMPUTERNAME - $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')"
            Send-DiagMail -Subject $subject -Body $script:LastReport
        }
    })

$buttonClose.Add_Click({
        $form.Close()
    })

# =========================
# UPDATE AUTO AU DEMARRAGE
# =========================
$startupTimer = New-Object System.Windows.Forms.Timer
$startupTimer.Interval = 800

$startupTimer.Add_Tick({
        $startupTimer.Stop()

        if (-not $script:UpdateChecked) {
            $script:UpdateChecked = $true
            Invoke-StartupUpdateCheck
        }
    })

$form.Add_Shown({
        $startupTimer.Start()
    })

# =========================
# AFFICHAGE
# =========================
[void]$form.ShowDialog()