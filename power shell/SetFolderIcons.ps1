# Lệnh trong PowerShell: Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass (bật quyền cho phép chạy mã trên hệ thống)
# Chuyển về đường dẫn chứa file SetFolderIcons.ps1, sau đó chạy .\SetFolderIcons.ps1
# Đường dẫn thư mục chính 
$rootFolder = "D:\AN THINH NAM\San_Pham\EuroPaint"
# Thư mục chứa các file ico
$iconFolder = "D:\AN THINH NAM\HINH ANH - VIDEO\02. NOI DUNG DU AN\08. ICON\icons_1"

# Lấy tất cả file .ico trong thư mục icon
Get-ChildItem -Path $iconFolder -Filter *.ico | ForEach-Object {
    $icoFile = $_.FullName
    $baseName = $_.BaseName  # tên không đuôi mở rộng

    # Tạo thư mục con nếu chưa có
    $subFolder = Join-Path $rootFolder $baseName
    if (-not (Test-Path $subFolder)) {
        New-Item -Path $subFolder -ItemType Directory | Out-Null
    }

    # Tạo file desktop.ini trong thư mục con
    $desktopIniPath = Join-Path $subFolder "desktop.ini"

    # Ghi thông tin vào desktop.ini
    @"
[.ShellClassInfo]
IconResource=$icoFile,0
"@ | Set-Content -Encoding Unicode -Path $desktopIniPath

    # Đặt thuộc tính cho folder và desktop.ini
    attrib +s "$subFolder"         # folder system để icon hiện
    attrib +h "$desktopIniPath"    # ẩn file desktop.ini
    attrib +r "$desktopIniPath"    # chỉ đọc

    Write-Host "✔ Đã tạo và gán icon cho thư mục: $baseName"
}
