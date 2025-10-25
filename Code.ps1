<# Windows PowerShell 5.1.22621.2506#>
# 彻底删除指定文件扩展名的所有注册表关联
# 以管理员身份运行

# 定义变量
$extensions = @(
    "avi", "wmv", "wmp", "wm", "asf", "mpg", "mpeg", "mpe", "m1v", "m2v",
    "mpv2", "mp2v", "ts", "tp", "tpr", "trp", "vob", "ifo", "ogm", "ogv",
    "mp4", "m4v", "m4p", "m4b", "3gp", "3gpp", "3g2", "3gp2", "mkv", "rm",
    "ram", "rmvb", "rpm", "flv", "mov", "qt", "nsv", "dpg", "m2ts"， "m2t",
    "mts", "dvr-ms", "k3g", "skm", "evo", "nsr", "amv", "divx", "webm"，
    "wtv", "f4v", "mxf", "wav", "wma", "mpa", "mp2", "m1a", "m2a", "mp3",
    "ogg", "m4a", "aac", "mka", "ra", "flac", "ape", "mpc", "mod", "ac3",
    "eac3", "dts", "dtshd", "wv", "tak", "cda", "dsf", "tta", "aiff", "aif",
    "opus", "amr", "asx", "m3u", "m3u8", "pls", "wvx", "wax", "wmx", "cue",
    "mpls", "mpl", "dpl", "xspf", "mpd"
)

# 初始化日志
$log = @()
$log += "===== Cleanup started: $(Get-Date) ====="

# 检查管理员权限
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "Please run this script as Administrator!"
    exit 1
}

# 删除注册表路径的函数
function Remove-RegistryPath {
    param([string]$Path)
    if (Test-Path $Path) {
        try {
            Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
            $script:log += "Successfully deleted: $Path"
        }
        catch {
            $script:log += "Warning: Failed to delete $Path. Error: $($_.Exception.Message)"
        }
    }
    else {
        $script:log += "Not found: $Path"
    }
}

# 处理每个扩展
foreach ($ext 在 $extensions) {
    $extWithDot = ".$ext"
    $log += "`n----- Processing: $extWithDot -----"

    # 1. 用户级 FileExts（包括手动关联）
    Remove-RegistryPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$extWithDot"

    # 2. 用户级类
    Remove-RegistryPath "HKCU:\Software\Classes\$extWithDot"

    # 3. 系统级类
    Remove-RegistryPath "HKLM:\SOFTWARE\Classes\$extWithDot"

    # 4. 清理关联的 ProgID
    $systemExtPath = "HKLM:\SOFTWARE\Classes\$extWithDot"
    if (Test-Path $systemExtPath) {
        $progId = (Get-ItemProperty $systemExtPath -ErrorAction SilentlyContinue).'(Default)'
        if ($progId) { Remove-RegistryPath "HKLM:\SOFTWARE\Classes\$progId" }
    }
    $userClassesPath = "HKCU:\Software\Classes\$extWithDot"
    if (Test-Path $userClassesPath) {
        $userProgId = (Get-ItemProperty $userClassesPath -ErrorAction SilentlyContinue).'(Default)'
        if ($userProgId) { Remove-RegistryPath "HKCU:\Software\Classes\$userProgId" }
    }
}

# 重新启动资源管理器以应用更改
try {
    $log += "`n----- Restarting Explorer -----"
    Stop-Process -Name explorer -Force -ErrorAction Stop
    Start-Process explorer -ErrorAction Stop
    $log += "Explorer restarted successfully"
}
catch {
    $log += "Warning: Failed to restart Explorer. Please restart your PC manually."
}

# 将日志保存到桌面
$log += "`n===== Cleanup completed: $(Get-Date) ====="
$logPath = "$env:USERPROFILE\Desktop\FileAssociationCleanupLog.txt"
$log | Out-File -FilePath $logPath -Encoding UTF8
Write-Host "`nOperation completed! Log saved to: $logPath`n"
