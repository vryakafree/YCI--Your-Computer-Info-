@echo off
setlocal EnableDelayedExpansion

rem === Ensure UTF-8 for box-drawing characters ===
chcp 65001 >nul

rem === Capture the ESC character into %ESC% ===
for /F "delims=" %%A in ('echo prompt $E^| cmd') do set "ESC=%%A"

rem === Optional: enable Virtual Terminal processing via registry (no admin needed) ===
rem This helps classic conhost honor ANSI sequences on newer Windows.
reg add HKCU\Console /v VirtualTerminalLevel /t REG_DWORD /d 1 /f >nul 2>&1

rem === Define color sequences ===
set "RST=%ESC%[0m"
set "BLUE=%ESC%[34m"
set "YEL=%ESC%[33m"

rem 256-color brown approximation (falls back visually to yellow if not supported)
set "BROWN256=%ESC%[38;5;94m"

cls

rem ====== Box Header ======
echo  ╔═════════════════════════════════════════════════════════════════════════╗
echo  ║                                                                         ║
echo  ║                         ██╗  ██╗ █████╗██████╗                          ║
echo  ║                         ██║  ██║██╔═══╝  ██╔═╝                          ║
echo  ║                         ╚██ ██╔╝██║      ██║                            ║
echo  ║                          ╚██╔═╝ ██║      ██║                            ║
echo  ║                           ██║   ╚█████╗██████╗                          ║
echo  ║                           ╚═╝    ╚════╝╚═════╝                          ║
echo  ║                                                                         ║
echo  ║                                                                         ║
echo  ╚═════════════════════════════════════════════════════════════════════════╝

echo(
echo                                 ──  Y C I  ──
echo                               Your Computer Info
echo(

rem === Small info lines with colors ===
rem version: "1.0" in brown (256-color), resets after
echo  version !BROWN256!1.0!RST!

rem Author line:
rem "Copilot, Gemini" in blue; "DiEn" in yellow; escape the ampersand with ^
echo  Author: !BLUE!Copilot!RST!, !BLUE!Gemini!RST! ^& !YEL!DiEn!RST!

:: --- THIET LAP FILE XUAT ---
set "FileName=YOUR_FILE_NAME.csv"
set "OutPath=%~dp0%FileName%"

echo Dang quet thong tin he thong...

:CheckFile
:: Kiem tra xem file co dang bi khoa (dang mo) khong
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$file = New-Object System.IO.FileInfo '%OutPath%';" ^
    "if (Test-Path '%OutPath%') {" ^
    "  try { $stream = $file.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None); $stream.Close() } " ^
    "  catch { exit 1 }" ^
    "}"
    
if %errorlevel% neq 0 (
    echo.
    echo [CANH BAO] File %FileName% dang mo!
    echo Vui long LUU va DONG file Excel lai, sau do quay lai day.
    echo.
    pause
    goto CheckFile
)

:: --- CHAY POWERSHELL QUET DU LIEU ---
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference = 'SilentlyContinue';" ^
    "function ToNA($v){ if($null -eq $v -or [string]::IsNullOrWhiteSpace($v)){ return 'N/A' } return $v.ToString().Trim() };" ^
    "$comName = $env:COMPUTERNAME;" ^
    "$sysInfo = Get-CimInstance Win32_ComputerSystem;" ^
    "$realUser = $sysInfo.UserName;" ^
    "if (-not $realUser) { $realUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name + ' (Admin Mode)' };" ^
    "$physicalAdapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' -and $_.InterfaceDescription -notlike '*Virtual*' -and $_.InterfaceDescription -notlike '*Hyper-V*' -and $_.InterfaceAlias -notlike 'vEthernet*' -and $_.InterfaceDescription -notlike '*VMware*' -and $_.InterfaceDescription -notlike '*VirtualBox*' };" ^
    "$allIPs = Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike '*Loopback*' -and $_.IPAddress -notlike '169.*' -and $_.InterfaceIndex -in $physicalAdapters.InterfaceIndex };" ^
    "$ip = ($allIPs | Sort-Object @{Expression={if($_.InterfaceAlias -like '*Wi-Fi*'){0}else{1}}}).IPAddress | Select-Object -First 1;" ^
    "$ip = ToNA $ip;" ^
    "$sn = (Get-CimInstance Win32_BIOS).SerialNumber; $sn = ToNA $sn;" ^
    "$bb = Get-CimInstance Win32_BaseBoard;" ^
    "$mainboard = if ($bb) { ($bb.Manufacturer + ' ' + $bb.Product).Trim() } else { 'N/A' };" ^
    "$cpu = (Get-CimInstance Win32_Processor).Name; $cpu = ToNA $cpu;" ^
    "$ramInfo = Get-CimInstance Win32_PhysicalMemory;" ^
    "$totalRam = [math]::round(($ramInfo | Measure-Object -Property Capacity -Sum).Sum / 1GB, 0);" ^
    "$typeCode = if($ramInfo){ if ($ramInfo[0].SMBIOSMemoryType) { $ramInfo[0].SMBIOSMemoryType } else { $ramInfo[0].MemoryType } } else {0};" ^
    "$memType = switch($typeCode) { 20 {'DDR2'} 24 {'DDR3'} 26 {'DDR4'} 30 {'DDR4'} 34 {'DDR5'} default {'RAM'} };" ^
    "$ramCombined = if($totalRam -gt 0){ ('{0}GB {1}' -f $totalRam, $memType) } else { 'N/A' };" ^
    "$slotsUsed = if($ramInfo){ @($ramInfo).Count } else {0};" ^
    "$slotsMax = try { (Get-CimInstance Win32_PhysicalMemoryArray).MemoryDevices } catch { 0 };" ^
    "$ramSlotsCombined = ' used ' + ('{0}/{1}' -f $slotsUsed, $slotsMax);" ^
    "$disks = $null; try { $pd = Get-PhysicalDisk; if($pd){ $disks = ($pd | ForEach-Object { '{0} ({1}GB)' -f $_.FriendlyName, [math]::Round($_.Size/1e9,0) }) -join ' | ' } } catch {};" ^
    "if (-not $disks) { $dd = Get-CimInstance Win32_DiskDrive; if($dd){ $disks = ($dd | ForEach-Object { '{0} ({1}GB)' -f $_.Model, [math]::Round($_.Size/1e9,0) }) -join ' | ' } };" ^
    "$disks = ToNA $disks;" ^
    "$gpus = (Get-CimInstance Win32_VideoController | Select-Object -ExpandProperty Name | Sort-Object -Unique) -join ' | '; $gpus = ToNA $gpus;" ^
    "$osObj = Get-CimInstance Win32_OperatingSystem;" ^
    "$os = if ($osObj) { '{0} (Build {1})' -f $osObj.Caption.Trim(), $osObj.BuildNumber } else { 'N/A' };" ^
    "$monNames = @(); try { $wmiMon = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID; foreach($m in $wmiMon){ $name = [System.Text.Encoding]::ASCII.GetString(($m.UserFriendlyName | Where-Object {$_ -ne 0})).Trim(); if($name){$monNames += $name} } } catch{};" ^
    "$monInches = @(); try { $monBasic = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams; foreach($b in $monBasic){ $size = [math]::Round([math]::Sqrt([math]::Pow(($b.MaxHorizontalImageSize/2.54),2) + [math]::Pow(($b.MaxVerticalImageSize/2.54),2)), 0); $monInches += $size } } catch{};" ^
    "$displayList = @(); $maxIdx = [math]::Max($monNames.Count, $monInches.Count);" ^
    "for($i=0; $i -lt $maxIdx; $i++){ $n = if($i -lt $monNames.Count){ $monNames[$i] } else {'Monitor'}; $inc = if($i -lt $monInches.Count){ $monInches[$i] } else {$null}; if($inc -gt 0){ $displayList += ( '{0}inch {1}' -f $inc, $n ) } else { $displayList += $n } };" ^
    "$displayCombined = if($displayList){ $displayList -join ' | ' } else { 'N/A' };" ^
    "$timeNow = Get-Date -Format 'dd/MM/yyyy HH:mm:ss';" ^
    "$columns = 'STT','User','ComputerName','IPAddress','SerialNumber','Mainboard','CPU','RAM','RAM_Slots','GPU','OS','Disks','Display','Logtimes';" ^
    "[array]$list = @();" ^
    "if (Test-Path '%OutPath%') { $list = @(Import-Csv '%OutPath%' -Encoding UTF8) };" ^
    "if ($list.Count -gt 0 -and $list[0].PSObject.Properties['Monitor']) { $list = @() };" ^
    "$existing = $list | Where-Object { $_.ComputerName -eq $comName };" ^
    "$props = [ordered]@{ STT=''; User=$realUser; ComputerName=$comName; IPAddress=$ip; SerialNumber=$sn; Mainboard=$mainboard; CPU=$cpu; RAM=$ramCombined; RAM_Slots=$ramSlotsCombined; GPU=$gpus; OS=$os; Disks=$disks; Display=$displayCombined; Logtimes=$timeNow };" ^
    "if ($existing) {" ^
    "   foreach($k in $props.Keys) { $existing[0].$k = $props[$k] }" ^
    "} else {" ^
    "   $list += [PSCustomObject]$props" ^
    "};" ^
    "$i=1; $sortedList = $list | Sort-Object ComputerName | Select-Object $columns; $sortedList | ForEach-Object { $_.STT = $i; $i++ };" ^
    "$sortedList | Export-Csv -Path '%OutPath%' -NoTypeInformation -Encoding UTF8 -Force"

if exist "%OutPath%" (
    echo.
    echo ====================================================
    echo HOAN THANH! Du lieu may %COMPUTERNAME% da duoc cap nhat.
    echo ====================================================
) else (
    echo.
    echo LOI: Khong the ghi du lieu.
)


pause
