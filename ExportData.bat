@echo off
setlocal EnableDelayedExpansion

:: --- THIẾT LẬP FILE XUẤT ---
set "FileName=%COMPUTERNAME%.txt"
set "OutPath=%~dp0%FileName%"

echo Dang quet thong tin he thong...
echo File se duoc luu tai: "%OutPath%"

:: --- CHẠY POWERSHELL VÀ XUẤT THẲNG RA FILE ---
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$sysInfo = Get-CimInstance Win32_ComputerSystem;" ^
    "$realUser = $sysInfo.UserName;" ^
    "if (-not $realUser) { $realUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name + ' (Admin Mode)' };" ^
    "$comName = $env:COMPUTERNAME;" ^
    "$sn = (Get-CimInstance Win32_BIOS).SerialNumber;" ^
    "$bb = Get-CimInstance Win32_BaseBoard;" ^
    "$mMain = $bb.Product;" ^
    "$mFac = $bb.Manufacturer;" ^
    "$cpu = (Get-CimInstance Win32_Processor).Name.Trim();" ^
    "$ramInfo = Get-CimInstance Win32_PhysicalMemory;" ^
    "$totalRam = ($ramInfo | Measure-Object -Property Capacity -Sum).Sum / 1GB;" ^
    "$finalRam = '' + [math]::round($totalRam, 2) + ' GB';" ^
    "$typeCode = $ramInfo[0].MemoryType;" ^
    "if ($typeCode -eq 0) { $typeCode = $ramInfo[0].SMBIOSMemoryType };" ^
    "$memType = switch($typeCode) { 20 {'DDR2'} 21 {'DDR2'} 24 {'DDR3'} 26 {'DDR4'} 30 {'LPDDR4'} 34 {'DDR5'} default {'Khong xac dinh'} };" ^
    "$disks = Get-PhysicalDisk | ForEach-Object { '{0} ({1} GB, {2})' -f $_.FriendlyName, [math]::round($_.Size/1GB, 2), $_.MediaType };" ^
    "$disksString = $disks -join ' | ';" ^
    "$monitors = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID | ForEach-Object {" ^
    "  $name = [System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName -ne 0).Trim();" ^
    "  if (-not $name) { $name = 'Unknown Monitor' }; $name" ^
    "};" ^
    "$monitorsInch = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams | ForEach-Object { $inch = [math]::Round([math]::Sqrt([math]::Pow($_.MaxHorizontalImageSize/2.54,2) + [math]::Pow($_.MaxVerticalImageSize/2.54,2)),1); Write-Output ('' + $inch + ' inch') };" ^
    "$monString = $monitors -join ' | ';" ^
    "$output = @();" ^
    "$output += '--- THONG TIN HE THONG ---';" ^
    "$output += '--- Ngay quet: ' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + ' ---';" ^
    "$output += 'Nguoi dung: ' + $realUser;" ^
    "$output += 'Ten may:    ' + $comName;" ^
    "$output += 'So Serial:  ' + $sn;" ^
    "$output += 'Mainboard:  ' + $mFac + ' ' + $mMain;" ^
    "$output += 'CPU:        ' + $cpu;" ^
    "$output += 'RAM:        ' + $finalRam + ' (' + $memType + ')';" ^
    "$output += 'O cung:     ' + $disksString;" ^
    "$output += 'Man hinh:   ' + $monString+ ' (' + $monitorsInch + ')';" ^
    "$output | Out-File -FilePath '%OutPath%' -Encoding UTF8 -Force"

if exist "%OutPath%" (
    echo.
    echo ====================================================
    echo THANH CONG! File da duoc tao tai:
    echo "%OutPath%"
    echo ====================================================
) else (
    echo.
    echo LOI: Khong the tao duoc file. Vui long kiem tra quyen ghi thu muc.
)

pause