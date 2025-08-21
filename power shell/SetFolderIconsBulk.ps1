# ===========================================
# SCRIPT TẠO FOLDER ICON HÀNG LOẠT CHO NHIỀU THƯ MỤC GỐC
# Lệnh trong PowerShell: Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass (bật quyền cho phép chạy mã trên hệ thống)
# ===========================================

param(
    [string[]]$RootFolders = @(),
    [string]$IconFolder = "",
    [switch]$ShowProgress = $true,
    [switch]$Interactive = $false
)

# Cấu hình mặc định
$defaultRootFolders = @(
    "D:\AN THINH NAM\San_Pham\Antech",
	"D:\AN THINH NAM\San_Pham\AsiaMortar",
	"D:\AN THINH NAM\San_Pham\AsiaStar",
	"D:\AN THINH NAM\San_Pham\Bautek",
	"D:\AN THINH NAM\San_Pham\Bestmix",
	"D:\AN THINH NAM\San_Pham\Bitumax",
	"D:\AN THINH NAM\San_Pham\BituNil",
	"D:\AN THINH NAM\San_Pham\Bossil",
	"D:\AN THINH NAM\San_Pham\EuroPaint",
	"D:\AN THINH NAM\San_Pham\Fosroc",
	"D:\AN THINH NAM\San_Pham\Intoc",
	"D:\AN THINH NAM\San_Pham\Lemax",
	"D:\AN THINH NAM\San_Pham\Mapei",
	"D:\AN THINH NAM\San_Pham\Modern",
	"D:\AN THINH NAM\San_Pham\Neomax",
	"D:\AN THINH NAM\San_Pham\Neotex",
	"D:\AN THINH NAM\San_Pham\Quicseal",
	"D:\AN THINH NAM\San_Pham\Shell",
	"D:\AN THINH NAM\San_Pham\SieuThiChongTham",
	"D:\AN THINH NAM\San_Pham\Sika",
	"D:\AN THINH NAM\San_Pham\Technonicol",
	"D:\AN THINH NAM\San_Pham\Victory"
    # Thêm các thư mục gốc khác vào đây
)

$defaultIconFolder = "D:\AN THINH NAM\HINH ANH - VIDEO\02. NOI DUNG DU AN\08. ICON\icons_1"

# Sử dụng giá trị mặc định nếu không được cung cấp
if ($RootFolders.Count -eq 0) {
    $RootFolders = $defaultRootFolders
}

if (-not $IconFolder) {
    $IconFolder = $defaultIconFolder
}

# Function: Tạo hoặc cập nhật folder icon
function Set-FolderIcon {
    param(
        [string]$FolderPath,
        [string]$IconPath,
        [string]$FolderName,
        [bool]$IsNewFolder
    )
    
    try {
        $desktopIniPath = Join-Path $FolderPath "desktop.ini"
        
        # Kiểm tra và xử lý file desktop.ini hiện có
        if (Test-Path $desktopIniPath) {
            try {
                # Xóa thuộc tính bảo vệ của file cũ
                attrib -h -r -s "$desktopIniPath" 2>$null
                Remove-Item -Path $desktopIniPath -Force -ErrorAction SilentlyContinue
                Write-Host "    🔄 Đã xóa desktop.ini cũ: $FolderName" -ForegroundColor Yellow
            }
            catch {
                Write-Host "    ⚠ Không thể xóa desktop.ini cũ: $FolderName" -ForegroundColor Yellow
            }
        }
        
        # Nội dung desktop.ini mới
        $desktopIniContent = @"
[.ShellClassInfo]
IconResource=$IconPath,0
"@
        
        # Thử tạo file desktop.ini mới
        try {
            # Ghi file với quyền cao nhất
            Set-Content -Path $desktopIniPath -Value $desktopIniContent -Encoding Unicode -Force
            
            # Đặt thuộc tính cho folder và desktop.ini
            Start-Process -FilePath "attrib" -ArgumentList "+s `"$FolderPath`"" -NoNewWindow -Wait -ErrorAction SilentlyContinue
            Start-Process -FilePath "attrib" -ArgumentList "+h +r `"$desktopIniPath`"" -NoNewWindow -Wait -ErrorAction SilentlyContinue
            
            if ($IsNewFolder) {
                Write-Host "    ✔ Tạo mới: $FolderName" -ForegroundColor Green
            } else {
                Write-Host "    ✔ Cập nhật: $FolderName" -ForegroundColor Cyan
            }
            
            return $true
        }
        catch {
            # Nếu vẫn lỗi, thử phương pháp khác
            Write-Host "    ⚠ Thử phương pháp khác cho: $FolderName" -ForegroundColor Yellow
            
            # Tạo file tạm thời
            $tempFile = Join-Path $env:TEMP "desktop_temp.ini"
            Set-Content -Path $tempFile -Value $desktopIniContent -Encoding Unicode
            
            # Copy file tạm thời vào thư mục đích
            Copy-Item -Path $tempFile -Destination $desktopIniPath -Force
            Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
            
            # Đặt thuộc tính
            Start-Process -FilePath "attrib" -ArgumentList "+s `"$FolderPath`"" -NoNewWindow -Wait -ErrorAction SilentlyContinue
            Start-Process -FilePath "attrib" -ArgumentList "+h +r `"$desktopIniPath`"" -NoNewWindow -Wait -ErrorAction SilentlyContinue
            
            Write-Host "    ✔ Phương pháp 2 thành công: $FolderName" -ForegroundColor Green
            return $true
        }
        
    }
    catch {
        Write-Host "    ✖ Lỗi quyền truy cập: $FolderName - $_" -ForegroundColor Red
        Write-Host "    💡 Gợi ý: Chạy PowerShell với quyền Administrator" -ForegroundColor Yellow
        return $false
    }
}

# Function: Xử lý một thư mục gốc
function Process-RootFolder {
    param(
        [string]$RootPath,
        [array]$IconFiles,
        [int]$RootIndex,
        [int]$TotalRoots
    )
    
    Write-Host ""
    Write-Host "[$($RootIndex + 1)/$TotalRoots] 📁 $RootPath" -ForegroundColor Yellow
    Write-Host "=" * 60 -ForegroundColor Gray
    
    # Kiểm tra và tạo thư mục gốc nếu cần
    if (-not (Test-Path $RootPath)) {
        Write-Host "  ⚠ Thư mục không tồn tại, đang tạo..." -ForegroundColor Yellow
        try {
            New-Item -Path $RootPath -ItemType Directory -Force | Out-Null
            Write-Host "  ✔ Đã tạo thư mục gốc" -ForegroundColor Green
        }
        catch {
            Write-Host "  ✖ Không thể tạo thư mục gốc: $_" -ForegroundColor Red
            return @{ New = 0; Updated = 0; Errors = 1 }
        }
    }
    
    # Bộ đếm cho thư mục gốc này
    $stats = @{ New = 0; Updated = 0; Errors = 0 }
    
    # Xử lý từng file icon
    for ($i = 0; $i -lt $IconFiles.Count; $i++) {
        $iconFile = $IconFiles[$i]
        
        if ($ShowProgress) {
            $percentComplete = [math]::Round(($i / $IconFiles.Count) * 100, 1)
            Write-Progress -Id 1 -Activity "Xử lý thư mục: $RootPath" -Status "($($i + 1)/$($IconFiles.Count)) $($iconFile.Name)" -PercentComplete $percentComplete
        }
        
        # Lấy tên folder từ tên file
        $folderName = $iconFile.BaseName
        $folderPath = Join-Path $RootPath $folderName
        
        # Kiểm tra folder có tồn tại không
        $isNewFolder = $false
        if (-not (Test-Path $folderPath)) {
            try {
                New-Item -Path $folderPath -ItemType Directory | Out-Null
                $isNewFolder = $true
            }
            catch {
                Write-Host "    ✖ Không thể tạo: $folderName - $_" -ForegroundColor Red
                $stats.Errors++
                continue
            }
        }
        
        # Tạo hoặc cập nhật icon
        $success = Set-FolderIcon -FolderPath $folderPath -IconPath $iconFile.FullName -FolderName $folderName -IsNewFolder $isNewFolder
        
        if ($success) {
            if ($isNewFolder) {
                $stats.New++
            } else {
                $stats.Updated++
            }
        } else {
            $stats.Errors++
        }
    }
    
    # Xóa progress bar cho thư mục này
    if ($ShowProgress) {
        Write-Progress -Id 1 -Activity "Xử lý thư mục: $RootPath" -Completed
    }
    
    # Hiển thị thống kê cho thư mục này
    Write-Host ""
    Write-Host "  📊 Kết quả cho thư mục này:" -ForegroundColor White
    Write-Host "    ✔ Tạo mới: $($stats.New)" -ForegroundColor Green
    Write-Host "    ✔ Cập nhật: $($stats.Updated)" -ForegroundColor Cyan
    if ($stats.Errors -gt 0) {
        Write-Host "    ✖ Lỗi: $($stats.Errors)" -ForegroundColor Red
    }
    
    return $stats
}

# Function: Chọn thư mục gốc tương tác
function Select-RootFolders {
    Write-Host ""
    Write-Host "=== CHỌN THƯ MỤC GỐC ===" -ForegroundColor Cyan
    Write-Host "Danh sách thư mục gốc có sẵn:" -ForegroundColor White
    
    for ($i = 0; $i -lt $RootFolders.Count; $i++) {
        Write-Host "$($i + 1). $($RootFolders[$i])" -ForegroundColor Yellow
    }
    
    Write-Host "A. Tất cả thư mục" -ForegroundColor Green
    Write-Host "C. Tùy chỉnh đường dẫn mới" -ForegroundColor Magenta
    Write-Host ""
    
    $selection = Read-Host "Chọn thư mục (số, A cho tất cả, C cho tùy chỉnh)"
    
    if ($selection -eq "A" -or $selection -eq "a") {
        return $RootFolders
    }
    elseif ($selection -eq "C" -or $selection -eq "c") {
        $customPaths = @()
        do {
            $customPath = Read-Host "Nhập đường dẫn thư mục (Enter để kết thúc)"
            if ($customPath -and $customPath.Trim() -ne "") {
                $customPaths += $customPath.Trim()
                Write-Host "✔ Đã thêm: $customPath" -ForegroundColor Green
            }
        } while ($customPath -and $customPath.Trim() -ne "")
        
        return $customPaths
    }
    else {
        try {
            $indices = $selection -split "," | ForEach-Object { [int]$_.Trim() - 1 }
            $selectedFolders = @()
            foreach ($index in $indices) {
                if ($index -ge 0 -and $index -lt $RootFolders.Count) {
                    $selectedFolders += $RootFolders[$index]
                }
            }
            
            if ($selectedFolders.Count -gt 0) {
                return $selectedFolders
            } else {
                Write-Host "Lựa chọn không hợp lệ!" -ForegroundColor Red
                return Select-RootFolders
            }
        }
        catch {
            Write-Host "Vui lòng nhập số hợp lệ hoặc A/C!" -ForegroundColor Red
            return Select-RootFolders
        }
    }
}

# Function: Kiểm tra quyền Administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function: Yêu cầu chạy với quyền Administrator
function Request-AdminRights {
    Write-Host ""
    Write-Host "⚠️  CẢNH BÁO: Không có quyền Administrator!" -ForegroundColor Yellow
    Write-Host "Một số thao tác có thể thất bại do hạn chế quyền truy cập." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "💡 Khuyến nghị:" -ForegroundColor Cyan
    Write-Host "   1. Đóng PowerShell hiện tại" -ForegroundColor White
    Write-Host "   2. Chuột phải vào PowerShell → 'Run as Administrator'" -ForegroundColor White
    Write-Host "   3. Chạy lại script này" -ForegroundColor White
    Write-Host ""
    
    $continue = Read-Host "Bạn có muốn tiếp tục không? (y/N)"
    if ($continue -ne "y" -and $continue -ne "Y") {
        Write-Host "Đã dừng script." -ForegroundColor Yellow
        exit 0
    }
}

# ===========================================
# MAIN SCRIPT
# ===========================================

Write-Host "=== TẠO FOLDER ICON HÀNG LOẠT CHO NHIỀU THƯ MỤC ===" -ForegroundColor Magenta

# Kiểm tra quyền Administrator
if (-not (Test-Administrator)) {
    Request-AdminRights
}
Write-Host "Thư mục chứa icon: $IconFolder" -ForegroundColor White

# Function: Sửa lỗi quyền truy cập hàng loạt
function Fix-PermissionIssues {
    param([array]$RootFolders)
    
    Write-Host ""
    Write-Host "🔧 CÔNG CỤ SỬA LỖI QUYỀN TRUY CẬP" -ForegroundColor Cyan
    Write-Host "Đang quét và sửa lỗi desktop.ini trong các thư mục..." -ForegroundColor White
    
    $fixedCount = 0
    $totalProcessed = 0
    
    foreach ($rootFolder in $RootFolders) {
        if (-not (Test-Path $rootFolder)) { continue }
        
        Write-Host ""
        Write-Host "📁 Đang xử lý: $rootFolder" -ForegroundColor Yellow
        
        $subFolders = Get-ChildItem -Path $rootFolder -Directory -ErrorAction SilentlyContinue
        
        foreach ($folder in $subFolders) {
            $totalProcessed++
            $desktopIniPath = Join-Path $folder.FullName "desktop.ini"
            
            if (Test-Path $desktopIniPath) {
                try {
                    # Xóa tất cả thuộc tính bảo vệ
                    Start-Process -FilePath "attrib" -ArgumentList "-h -r -s `"$desktopIniPath`"" -NoNewWindow -Wait -ErrorAction SilentlyContinue
                    
                    # Đặt quyền full control
                    $acl = Get-Acl $desktopIniPath
                    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($env:USERNAME, "FullControl", "Allow")
                    $acl.SetAccessRule($accessRule)
                    Set-Acl -Path $desktopIniPath -AclObject $acl -ErrorAction SilentlyContinue
                    
                    Write-Host "  ✔ Đã sửa: $($folder.Name)" -ForegroundColor Green
                    $fixedCount++
                }
                catch {
                    Write-Host "  ✖ Không sửa được: $($folder.Name)" -ForegroundColor Red
                }
            }
        }
    }
    
    Write-Host ""
    Write-Host "📊 Kết quả sửa lỗi:" -ForegroundColor White
    Write-Host "  ✔ Đã sửa: $fixedCount" -ForegroundColor Green
    Write-Host "  📄 Tổng quét: $totalProcessed" -ForegroundColor White
    
    if ($fixedCount -gt 0) {
        Write-Host ""
        Write-Host "💡 Đã sửa xong! Bạn có thể chạy lại script chính." -ForegroundColor Cyan
    }
}

Write-Host ""
Write-Host "Sẽ xử lý $($RootFolders.Count) thư mục gốc:" -ForegroundColor White
foreach ($root in $RootFolders) {
    Write-Host "  📁 $root" -ForegroundColor Gray
}

# Kiểm tra thư mục icon
if (-not (Test-Path $IconFolder)) {
    Write-Host ""
    Write-Host "✖ Thư mục chứa icon không tồn tại: $IconFolder" -ForegroundColor Red
    exit 1
}

# Lấy tất cả file .ico
$iconFiles = Get-ChildItem -Path $IconFolder -Filter "*.ico" -File

if ($iconFiles.Count -eq 0) {
    Write-Host ""
    Write-Host "✖ Không tìm thấy file .ico nào trong thư mục: $IconFolder" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Tìm thấy $($iconFiles.Count) file icon (.ico)" -ForegroundColor White

# Xác nhận trước khi bắt đầu
if ($Interactive) {
    $confirm = Read-Host "`nBạn có muốn tiếp tục? (y/N)"
    if ($confirm -ne "y" -and $confirm -ne "Y") {
        Write-Host "Đã hủy thao tác." -ForegroundColor Yellow
        exit 0
    }
}

# Bộ đếm tổng
$totalStats = @{ New = 0; Updated = 0; Errors = 0 }

# Progress bar tổng
if ($ShowProgress) {
    Write-Progress -Id 0 -Activity "Xử lý tất cả thư mục gốc" -Status "Bắt đầu..." -PercentComplete 0
}

# Xử lý từng thư mục gốc
for ($rootIndex = 0; $rootIndex -lt $RootFolders.Count; $rootIndex++) {
    $rootFolder = $RootFolders[$rootIndex]
    
    if ($ShowProgress) {
        $percentComplete = [math]::Round(($rootIndex / $RootFolders.Count) * 100, 1)
        Write-Progress -Id 0 -Activity "Xử lý tất cả thư mục gốc" -Status "($($rootIndex + 1)/$($RootFolders.Count)) $rootFolder" -PercentComplete $percentComplete
    }
    
    # Xử lý thư mục gốc này
    $stats = Process-RootFolder -RootPath $rootFolder -IconFiles $iconFiles -RootIndex $rootIndex -TotalRoots $RootFolders.Count
    
    # Cộng dồn thống kê
    $totalStats.New += $stats.New
    $totalStats.Updated += $stats.Updated
    $totalStats.Errors += $stats.Errors
}

# Xóa progress bar tổng
if ($ShowProgress) {
    Write-Progress -Id 0 -Activity "Xử lý tất cả thư mục gốc" -Completed
}

# Hiển thị thống kê tổng
Write-Host ""
Write-Host "=" * 80 -ForegroundColor Gray
Write-Host "=== THỐNG KÊ TỔNG KẾT ===" -ForegroundColor Magenta
Write-Host "🏗️  Thư mục gốc đã xử lý: $($RootFolders.Count)" -ForegroundColor White
Write-Host "📄 File icon đã quét: $($iconFiles.Count)" -ForegroundColor White
Write-Host "✔️  Thư mục mới được tạo: $($totalStats.New)" -ForegroundColor Green
Write-Host "🔄 Thư mục được cập nhật: $($totalStats.Updated)" -ForegroundColor Cyan
if ($totalStats.Errors -gt 0) {
    Write-Host "❌ Lỗi xảy ra: $($totalStats.Errors)" -ForegroundColor Red
}
Write-Host "📊 Tổng thư mục đã xử lý: $($totalStats.New + $totalStats.Updated)" -ForegroundColor Yellow

# Hỏi có muốn mở thư mục không
if ($Interactive -and $RootFolders.Count -le 3) {
    Write-Host ""
    Write-Host "Bạn có muốn mở các thư mục kết quả không? (y/N)" -ForegroundColor Yellow
    $openFolders = Read-Host
    if ($openFolders -eq "y" -or $openFolders -eq "Y") {
        foreach ($rootFolder in $RootFolders) {
            if (Test-Path $rootFolder) {
                try {
                    Start-Process explorer.exe -ArgumentList $rootFolder
                }
                catch {
                    Write-Host "Không thể mở thư mục: $rootFolder" -ForegroundColor Red
                }
            }
        }
    }
}

Write-Host ""
Write-Host "=== HOÀN THÀNH TẤT CẢ ===" -ForegroundColor Green
