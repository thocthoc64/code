# Script tạo shortcut hàng loạt + gán icon theo icon của THƯƠNG HIỆU (đọc từ desktop.ini)
# PowerShell 5+ | Chạy với quyền người dùng thường là đủ
# Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass (bật quyền cho phép chạy script trên hệ thống)

# ====== CẤU HÌNH ======
$rootPath   = "D:\AN THINH NAM\San_Pham"    # thư mục gốc chứa các thương hiệu
$tongHopPath = Join-Path $rootPath "00_TONG_HOP"

# ====== HÀM TIỆN ÍCH ======
function Get-FolderIconInfo {
    param(
        [Parameter(Mandatory)] [string] $FolderPath
    )
    $desktopIni = Join-Path $FolderPath "desktop.ini"
    if (-not (Test-Path $desktopIni)) { return $null }

    # Đọc ini
    $lines = Get-Content -LiteralPath $desktopIni -ErrorAction SilentlyContinue
    if (-not $lines) { return $null }

    # Lấy section [.ShellClassInfo]
    $sectionStart = ($lines | Select-String '^\s*\[\.ShellClassInfo\]\s*$').LineNumber
    if (-not $sectionStart) { return $null }
    $rest = $lines[$sectionStart..($lines.Count-1)]

    # Tạo bảng key=value
    $kv = @{}
    foreach ($ln in $rest) {
        if ($ln -match '^\s*\[') { break }             # hết section
        if ($ln -match '^\s*([A-Za-z0-9_]+)\s*=\s*(.*)\s*$') {
            $kv[$Matches[1]] = $Matches[2].Trim('"')
        }
    }

    # Ưu tiên IconResource (dạng "path,index")
    if ($kv.ContainsKey('IconResource') -and $kv['IconResource']) {
        $val = $kv['IconResource']
        # tách path và index (nếu có)
        $path, $idx = $val -split '\s*,\s*', 2
        if (-not $idx) { $idx = '0' }

        # Nếu path là tương đối (.\icon.ico hoặc icon.ico) => đổi sang tuyệt đối
        if ($path -match '^[\.\\]' -or -not ($path -match '^[A-Za-z]:\\' -or $path -match '^[\\/]')) {
            $path = Join-Path $FolderPath $path
        }
        return @{ Path = $path; Index = [int]$idx }
    }

    # Hoặc IconFile + IconIndex
    if ($kv.ContainsKey('IconFile') -and $kv['IconFile']) {
        $path = $kv['IconFile']
        $idx  = 0
        if ($kv.ContainsKey('IconIndex') -and $kv['IconIndex']) {
            $idx = [int]$kv['IconIndex']
        }
        if ($path -match '^[\.\\]' -or -not ($path -match '^[A-Za-z]:\\' -or $path -match '^[\\/]')) {
            $path = Join-Path $FolderPath $path
        }
        return @{ Path = $path; Index = [int]$idx }
    }

    # Không có cấu hình rõ ràng -> thử tìm .ico trong thư mục
    $ico = Get-ChildItem -LiteralPath $FolderPath -Filter *.ico -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($ico) { return @{ Path = $ico.FullName; Index = 0 } }

    return $null
}

# ====== CHUẨN BỊ THƯ MỤC ======
if (!(Test-Path $tongHopPath)) {
    New-Item -ItemType Directory -Path $tongHopPath -Force | Out-Null
    Write-Host "Đã tạo thư mục: $tongHopPath" -ForegroundColor Green
}

# Danh sách thương hiệu (loại trừ 00_TONG_HOP)
$brandFolders = Get-ChildItem -Path $rootPath -Directory | Where-Object { $_.Name -ne "00_TONG_HOP" }
if (-not $brandFolders -or $brandFolders.Count -eq 0) {
    Write-Host "Không tìm thấy folder thương hiệu nào!" -ForegroundColor Red
    exit
}

# Tập hợp tên các thư mục con duy nhất
$allSubfolders = @()
foreach ($brandFolder in $brandFolders) {
    $subfolders = Get-ChildItem -Path $brandFolder.FullName -Directory -ErrorAction SilentlyContinue
    foreach ($subfolder in $subfolders) {
        if ($allSubfolders -notcontains $subfolder.Name) { $allSubfolders += $subfolder.Name }
    }
}

# COM object tạo shortcut
$WshShell = New-Object -ComObject WScript.Shell

# Cache icon cho mỗi thương hiệu (đỡ phải đọc desktop.ini lặp lại)
$brandIconMap = @{}
foreach ($brand in $brandFolders) {
    $iconInfo = Get-FolderIconInfo -FolderPath $brand.FullName
    if ($iconInfo) {
        $brandIconMap[$brand.FullName] = "$($iconInfo.Path),$($iconInfo.Index)"
    } else {
        # fallback: để trống -> dùng icon mặc định của shortcut
        $brandIconMap[$brand.FullName] = $null
    }
}

# ====== TẠO SHORTCUT ======
foreach ($subfolderName in $allSubfolders) {
    $targetSubfolderPath = Join-Path $tongHopPath $subfolderName
    if (!(Test-Path $targetSubfolderPath)) {
        New-Item -ItemType Directory -Path $targetSubfolderPath -Force | Out-Null
        Write-Host "Đã tạo thư mục: $subfolderName" -ForegroundColor Green
    }

    foreach ($brandFolder in $brandFolders) {
        $sourceSubfolderPath = Join-Path $brandFolder.FullName $subfolderName
        if (Test-Path $sourceSubfolderPath) {
            $shortcutPath = Join-Path $targetSubfolderPath ("{0}.lnk" -f $brandFolder.Name)
            if (!(Test-Path $shortcutPath)) {
                $lnk = $WshShell.CreateShortcut($shortcutPath)
                $lnk.TargetPath = $sourceSubfolderPath
                $lnk.Description = "Shortcut to $($brandFolder.Name) - $subfolderName"
                $lnk.WorkingDirectory = $sourceSubfolderPath

                # >>> GÁN ICON TỪ THƯƠNG HIỆU <<<
                $iconLoc = $brandIconMap[$brandFolder.FullName]
                if ($iconLoc) {
                    $lnk.IconLocation = $iconLoc  # ví dụ "D:\...\brand.ico,0" hoặc "C:\Windows\System32\shell32.dll,4"
                }

                $lnk.Save()
                Write-Host "✓ Tạo shortcut + icon: $($brandFolder.Name) -> $subfolderName" -ForegroundColor Green
            } else {
                # Nếu đã tồn tại nhưng muốn cập nhật icon, bỏ comment 3 dòng dưới
                # $lnk = $WshShell.CreateShortcut($shortcutPath)
                # if ($brandIconMap[$brandFolder.FullName]) { $lnk.IconLocation = $brandIconMap[$brandFolder.FullName]; $lnk.Save() }
                Write-Host "- Shortcut đã tồn tại: $($brandFolder.Name) -> $subfolderName" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "⚠ Không tìm thấy: $($brandFolder.Name)\$subfolderName" -ForegroundColor Red
        }
    }
}

Write-Host "`n🎉 Hoàn thành! Shortcut đã được tạo và gán icon theo thương hiệu." -ForegroundColor Green
Write-Host "Nếu icon chưa đổi ngay, thử F5; hoặc chạy rebuild icon cache (mạnh tay) bằng cách restart Explorer." -ForegroundColor DarkYellow
