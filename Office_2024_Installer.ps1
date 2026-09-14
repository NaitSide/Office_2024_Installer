# Выборочная установка Microsoft 365 из IMG/ISO с помощью Office Deployment Tool.
$ErrorActionPreference = 'Stop'
$script:ImagePath = ''
$script:MountedByScript = $false
$script:OfficeProcess = $null
$script:ConfigPath = $null
$script:OdtPath = Join-Path $PSScriptRoot 'ODT\setup.exe'

$principal = [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    if (-not $PSCommandPath) { throw 'Сначала сохраните скрипт в файл.' }
    Start-Process powershell.exe -Verb RunAs -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',('"{0}"' -f $PSCommandPath))
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

function Add-Log([string]$Message) {
    $log.AppendText("$(Get-Date -Format HH:mm:ss)  $Message`r`n")
    $log.SelectionStart = $log.TextLength
    $log.ScrollToCaret()
}
function Show-Error([string]$Message) {
    Add-Log "Ошибка: $Message"
    [System.Windows.Forms.MessageBox]::Show($Message,'Ошибка','OK','Error') | Out-Null
}
function Get-ImageRoot {
    if (-not (Test-Path -LiteralPath $script:ImagePath -PathType Leaf)) { throw "Образ не найден: $script:ImagePath" }
    $image = Get-DiskImage -ImagePath $script:ImagePath
    if (-not $image.Attached) {
        Mount-DiskImage -ImagePath $script:ImagePath -ErrorAction Stop | Out-Null
        $script:MountedByScript = $true
    }
    $volume = Get-DiskImage -ImagePath $script:ImagePath | Get-Volume | Where-Object DriveLetter | Select-Object -First 1
    if (-not $volume) { throw 'У образа нет буквы диска.' }
    $root = "$($volume.DriveLetter):\"
    if (-not (Test-Path -LiteralPath (Join-Path $root 'Office\Data\C2RFireFlyData.xml'))) {
        throw 'В образе не найдены данные Office Click-to-Run.'
    }
    return $root
}
function Get-ImageInfo([string]$Root) {
    $dataRoot = Join-Path $Root 'Office\Data'
    $version = Get-ChildItem -LiteralPath $dataRoot -Directory |
        Where-Object { $_.Name -match '^16\.0\.\d+\.\d+$' } |
        Sort-Object { [version]$_.Name } -Descending |
        Select-Object -First 1 -ExpandProperty Name
    if (-not $version) { throw 'В образе не найдена версия Office.' }
    $folder = Join-Path $dataRoot $version
    $languages = @(Get-ChildItem -LiteralPath $folder -File -Filter 'stream.*.dat' |
        Where-Object { $_.Name -match '^stream\.(x64|x86)\.([a-z]{2}-[a-z]{2})\.dat$' } |
        ForEach-Object { [regex]::Match($_.Name,'^stream\.(x64|x86)\.([a-z]{2}-[a-z]{2})\.dat$').Groups[2].Value } |
        Sort-Object -Unique)
    if ($languages.Count -eq 0) { throw 'В образе не найдены языковые файлы Office.' }
    return [pscustomobject]@{ Version=$version; Folder=$folder; Languages=$languages }
}
function Cleanup-Install {
    if ($script:ConfigPath -and (Test-Path -LiteralPath $script:ConfigPath)) {
        Remove-Item -LiteralPath $script:ConfigPath -Force -ErrorAction SilentlyContinue
        $script:ConfigPath = $null
    }
    if ($script:MountedByScript) {
        try { Dismount-DiskImage -ImagePath $script:ImagePath -ErrorAction Stop | Out-Null; Add-Log 'Образ отключён.' }
        catch { Add-Log "Не удалось отключить образ: $($_.Exception.Message)" }
        $script:MountedByScript = $false
    }
}

$form = [System.Windows.Forms.Form]::new()
$form.Text = 'Установка Office 2024 из IMG'
$form.Size = [System.Drawing.Size]::new(520,545)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(30,30,40)
$form.ForeColor = [System.Drawing.Color]::White
$form.Font = [System.Drawing.Font]::new('Segoe UI',9)
function Add-Label([string]$Text,[int]$X,[int]$Y,[int]$Width,[int]$Height) {
    $control = [System.Windows.Forms.Label]::new()
    $control.Text = $Text
    $control.Location = [System.Drawing.Point]::new($X,$Y)
    $control.Size = [System.Drawing.Size]::new($Width,$Height)
    $form.Controls.Add($control)
    return $control
}
$title = Add-Label 'Выберите приложения Microsoft Office' 20 12 470 30
$title.Font = [System.Drawing.Font]::new('Segoe UI',13,[System.Drawing.FontStyle]::Bold)
$title.ForeColor = [System.Drawing.Color]::FromArgb(160,130,255)
$null = Add-Label 'Образ IMG:' 20 49 470 20
$imageBox = [System.Windows.Forms.TextBox]::new()
$imageBox.Location = [System.Drawing.Point]::new(20,69)
$imageBox.Size = [System.Drawing.Size]::new(390,26)
$form.Controls.Add($imageBox)
$browse = [System.Windows.Forms.Button]::new()
$browse.Text = 'Обзор'
$browse.Location = [System.Drawing.Point]::new(420,69)
$browse.Size = [System.Drawing.Size]::new(75,26)
$form.Controls.Add($browse)

$apps = [ordered]@{ Word=$true; Excel=$true; PowerPoint=$true; Outlook=$false; OneNote=$false; Access=$false; Publisher=$false; OneDrive=$false; Teams=$false }
$checks = @{}
$index = 0
foreach ($app in $apps.Keys) {
    $check = [System.Windows.Forms.CheckBox]::new()
    $check.Text = $app
    $check.Checked = $apps[$app]
    $check.ForeColor = [System.Drawing.Color]::White
    $check.Location = [System.Drawing.Point]::new((25 + [math]::Floor($index / 3) * 160),(108 + ($index % 3) * 29))
    $check.Size = [System.Drawing.Size]::new(145,25)
    $form.Controls.Add($check)
    $checks[$app] = $check
    $index++
}
$null = Add-Label 'Разрядность:' 20 205 100 25
$edition = [System.Windows.Forms.ComboBox]::new()
$edition.DropDownStyle = 'DropDownList'
$edition.Items.AddRange([object[]]@('64-бит','32-бит'))
$edition.SelectedIndex = 0
$edition.Location = [System.Drawing.Point]::new(120,202)
$edition.Size = [System.Drawing.Size]::new(95,28)
$form.Controls.Add($edition)
$languageLabel = Add-Label 'Язык: выберите образ' 240 205 255 25

$download = [System.Windows.Forms.LinkLabel]::new()
$download.Text = 'Скачать официальный Office'
$download.LinkColor = [System.Drawing.Color]::FromArgb(180,155,255)
$download.Location = [System.Drawing.Point]::new(20,240)
$download.Size = [System.Drawing.Size]::new(300,23)
$form.Controls.Add($download)
$download.Add_LinkClicked({ Start-Process 'https://massgrave.dev/office_c2r_links#russian-ru-ru' })
$null = Add-Label 'Russian [ru-RU] → вкладка «Office 2024»' 20 262 470 20
$null = Add-Label 'Product ID: ProPlus2024Retail → Offline x32–x64' 20 281 470 20

$odtInfo = Add-Label 'ODT — Office Deployment Tool: нужен для установки.' 20 310 475 20
$odtStatus = Add-Label '○ Проверяем ODT\setup.exe рядом со скриптом…' 20 331 475 22
$odtStatus.ForeColor = [System.Drawing.Color]::FromArgb(240,190,120)
$install = [System.Windows.Forms.Button]::new()
$install.Text = 'Установить'
$install.Enabled = $false
$install.Font = [System.Drawing.Font]::new('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$install.BackColor = [System.Drawing.Color]::FromArgb(100,70,200)
$install.ForeColor = [System.Drawing.Color]::White
$install.FlatStyle = 'Flat'
$install.FlatAppearance.BorderSize = 0
$install.Location = [System.Drawing.Point]::new(20,363)
$install.Size = [System.Drawing.Size]::new(150,36)
$form.Controls.Add($install)
$log = [System.Windows.Forms.TextBox]::new()
$log.Multiline = $true
$log.ReadOnly = $true
$log.ScrollBars = 'Vertical'
$log.Font = [System.Drawing.Font]::new('Consolas',9)
$log.BackColor = [System.Drawing.Color]::FromArgb(20,20,28)
$log.ForeColor = [System.Drawing.Color]::FromArgb(220,220,230)
$log.Location = [System.Drawing.Point]::new(20,411)
$log.Size = [System.Drawing.Size]::new(475,89)
$form.Controls.Add($log)

function Refresh-Readiness {
    $odtReady = Test-Path -LiteralPath $script:OdtPath -PathType Leaf
    if ($odtReady) {
        $odtStatus.Text = '✓ ODT\setup.exe найден рядом со скриптом'
        $odtStatus.ForeColor = [System.Drawing.Color]::FromArgb(130,220,150)
    } else {
        $odtStatus.Text = '○ Нет ODT\setup.exe рядом со скриптом'
        $odtStatus.ForeColor = [System.Drawing.Color]::FromArgb(240,190,120)
    }
    $candidate = $imageBox.Text.Trim()
    $imageReady = -not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate -PathType Leaf)
    $install.Enabled = $odtReady -and $imageReady -and (-not $script:OfficeProcess)
}
function Update-ImageLanguage {
    $candidate = $imageBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($candidate) -or -not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
        $languageLabel.Text = 'Язык: выберите образ'
        return
    }
    try {
        $script:ImagePath = $candidate
        $root = Get-ImageRoot
        $info = Get-ImageInfo $root
        $shown = @($info.Languages | ForEach-Object { if ($_ -eq 'ru-ru') { 'Русский (ru-ru)' } else { $_ } }) -join ', '
        $languageLabel.Text = "Язык образа: $shown"
        Add-Log "Образ: версия $($info.Version); язык: $shown"
    } catch { $languageLabel.Text = 'Язык: не определён'; Show-Error $_.Exception.Message }
    finally { Cleanup-Install }
}
$browse.Add_Click({
    $dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $dialog.Filter = 'Образы дисков (*.img;*.iso)|*.img;*.iso|Все файлы (*.*)|*.*'
    if ($dialog.ShowDialog() -eq 'OK') { $imageBox.Text = $dialog.FileName; Update-ImageLanguage }
    $dialog.Dispose()
})
$imageBox.Add_TextChanged({ Refresh-Readiness })
$imageBox.Add_Leave({ if (-not $script:OfficeProcess) { Update-ImageLanguage } })

$readinessTimer = [System.Windows.Forms.Timer]::new()
$readinessTimer.Interval = 1500
$readinessTimer.Add_Tick({ Refresh-Readiness })
$readinessTimer.Start()
$processTimer = [System.Windows.Forms.Timer]::new()
$processTimer.Interval = 1000
$processTimer.Add_Tick({
    try {
        if ($script:OfficeProcess -and $script:OfficeProcess.HasExited) {
            $processTimer.Stop()
            $exitCode = $script:OfficeProcess.ExitCode
            $script:OfficeProcess.Dispose()
            $script:OfficeProcess = $null
            Cleanup-Install
            Refresh-Readiness
            if ($exitCode -eq 0) { Add-Log 'Установка завершена успешно.'; [System.Windows.Forms.MessageBox]::Show('Установка завершена.','Готово','OK','Information') | Out-Null }
            else { Show-Error "Установщик завершился с кодом $exitCode. Журнал Office находится в TEMP." }
        }
    } catch { $processTimer.Stop(); Cleanup-Install; $script:OfficeProcess=$null; Refresh-Readiness; Show-Error $_.Exception.Message }
})
$install.Add_Click({
    try {
        Refresh-Readiness
        if (-not $install.Enabled) { throw 'Выберите существующий образ и проверьте ODT\setup.exe.' }
        $odtFile = Get-Item -LiteralPath $script:OdtPath
        $signature = Get-AuthenticodeSignature -LiteralPath $script:OdtPath
        if ($odtFile.VersionInfo.CompanyName -notlike 'Microsoft*' -or $signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notlike '*Microsoft Corporation*') {
            throw 'Файл ODT\setup.exe не имеет действительной подписи Microsoft.'
        }
        $script:ImagePath = $imageBox.Text.Trim()
        $selected = @($apps.Keys | Where-Object { $checks[$_].Checked })
        if ($selected.Count -eq 0) { throw 'Выберите хотя бы одно приложение.' }
        $root = Get-ImageRoot
        $info = Get-ImageInfo $root
        $arch = if ($edition.SelectedIndex -eq 1) { '32' } else { '64' }
        $streamArch = if ($arch -eq '32') { 'x86' } else { 'x64' }
        $available = @($info.Languages | Where-Object { Test-Path -LiteralPath (Join-Path $info.Folder "stream.$streamArch.$_.dat") })
        if ($available.Count -eq 0) { throw "В образе нет языка для $arch-битной установки." }
        $language = if ($available -contains 'ru-ru') { 'ru-ru' } elseif ($available.Count -eq 1) { $available[0] } else { throw "В образе несколько языков: $($available -join ', '). Нужен ru-ru." }
        $excluded = @($apps.Keys | Where-Object { -not $checks[$_].Checked })
        if (-not $checks['OneDrive'].Checked) { $excluded += 'Groove' }
        $excluded += 'Lync'                 # Skype для бизнеса — отдельный ID Office.
        $excluded += 'OutlookForWindows'    # Новый Outlook не выбирается в этом GUI.
        $excludeXml = ($excluded | Select-Object -Unique | ForEach-Object { "      <ExcludeApp ID=`"$_`" />" }) -join "`r`n"
        $sourcePath = [System.Security.SecurityElement]::Escape($root.TrimEnd('\'))
        $xml = @"
<Configuration>
  <Add SourcePath="$sourcePath" OfficeClientEdition="$arch" Version="$($info.Version)" AllowCdnFallback="False">
    <Product ID="ProPlus2024Retail">
      <Language ID="$language" />
$excludeXml
    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="FALSE" />
</Configuration>
"@
        $script:ConfigPath = Join-Path $env:TEMP ("OfficeSelection_{0}.xml" -f [guid]::NewGuid().ToString('N'))
        [System.IO.File]::WriteAllText($script:ConfigPath,$xml,[System.Text.UTF8Encoding]::new($false))
        $start = [System.Diagnostics.ProcessStartInfo]::new()
        $start.FileName = $script:OdtPath
        $start.Arguments = "/configure `"$script:ConfigPath`""
        $start.WorkingDirectory = Split-Path -Parent $script:OdtPath
        $start.UseShellExecute = $false
        $script:OfficeProcess = [System.Diagnostics.Process]::Start($start)
        if (-not $script:OfficeProcess) { throw 'Не удалось запустить ODT.' }
        Refresh-Readiness
        $processTimer.Start()
        Add-Log "Источник: $root | версия: $($info.Version) | $arch-бит | $language"
        Add-Log "Выбрано: $($selected -join ', ')"
        Add-Log "Исключено: $($excluded -join ', ')"
        Add-Log 'Установка запущена. Не закрывайте окно до завершения.'
    } catch { if (-not $script:OfficeProcess) { Cleanup-Install }; Refresh-Readiness; Show-Error $_.Exception.Message }
})
$form.Add_FormClosing({
    if ($script:OfficeProcess -and -not $script:OfficeProcess.HasExited) {
        $_.Cancel = $true
        [System.Windows.Forms.MessageBox]::Show('Дождитесь завершения установки Office.','Установка выполняется','OK','Information') | Out-Null
    }
})
Refresh-Readiness
Add-Log 'Готово к выбору приложений.'
[void]$form.ShowDialog()
$readinessTimer.Dispose()
$processTimer.Dispose()
$form.Dispose()
