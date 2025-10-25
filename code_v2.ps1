<# Windows PowerShell 5.1.22621.2506#>
# 彻底删除指定文件扩展名的所有注册表关联
# 以管理员身份运行

# 定义需要处理的文件扩展名
$extensions = @(
    "avi", "wmv", "wmp", "wm", "asf", "mpg", "mpeg", "mpe", "m1v", "m2v",
    "mpv2", "mp2v", "ts", "tp", "tpr", "trp", "vob", "ifo", "ogm", "ogv",
    "mp4", "m4v", "m4p", "m4b", "3gp", "3gpp", "3g2", "3gp2", "mkv", "rm",
    "ram", "rmvb", "rpm", "flv", "mov", "qt", "nsv", "dpg", "m2ts", "m2t",
    "mts", "dvr-ms", "k3g", "skm", "evo", "nsr", "amv", "divx", "webm",
    "wtv", "f4v", "mxf", "wav", "wma", "mpa", "mp2", "m1a", "m2a", "mp3",
    "ogg", "m4a", "aac", "mka", "ra", "flac", "ape", "mpc", "mod", "ac3",
    "eac3", "dts", "dtshd", "wv", "tak", "cda", "dsf", "tta", "aiff", "aif",
    "opus", "amr", "asx", "m3u", "m3u8", "pls", "wvx", "wax", "wmx", "cue",
    "mpls", "mpl", "dpl", "xspf", "mpd"
)

# 初始化日志数组
$log = @()
$startMsg = "===== Cleanup started: $(Get-Date) ====="
$log += $startMsg
Write-Host $startMsg  # 控制台显示开始信息

# 检查是否以管理员身份运行
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    $errorMsg = "Error: Please run this script as Administrator!"
    Write-Error $errorMsg
    $log += $errorMsg
    exit 1
}

# 删除注册表路径的函数（带控制台输出）
function Remove-RegistryPath {
    param([string]$Path)
    if (Test-Path $Path) {
        try {
            Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
            $msg = "Successfully deleted: $Path"
            $script:log += $msg
            Write-Host $msg  # 控制台显示成功信息
        }
        catch {
            $msg = "Warning: Failed to delete $Path. Error: $($_.Exception.Message)"
            $script:log += $msg
            Write-Host $msg  # 控制台显示错误信息
        }
    }
    else {
        $msg = "Not found: $Path"
        $script:log += $msg
        Write-Host $msg  # 控制台显示未找到信息
    }
}

# 逐个处理每个文件扩展名
foreach ($ext in $extensions) {
    $extWithDot = ".$ext"
    $processMsg = "`n----- Processing: $extWithDot -----"
    $log += $processMsg
    Write-Host $processMsg  # 控制台显示当前处理的扩展名

    # 1. 删除用户级FileExts关联
    Remove-RegistryPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$extWithDot"

    # 2. 删除用户级类关联
    Remove-RegistryPath "HKCU:\Software\Classes\$extWithDot"

    # 3. 删除系统级类关联
    Remove-RegistryPath "HKLM:\SOFTWARE\Classes\$extWithDot"

    # 4. 清理关联的ProgID（系统级）
    $systemExtPath = "HKLM:\SOFTWARE\Classes\$extWithDot"
    if (Test-Path $systemExtPath) {
        $progId = (Get-ItemProperty $systemExtPath -ErrorAction SilentlyContinue).'(Default)'
        if ($progId) { Remove-RegistryPath "HKLM:\SOFTWARE\Classes\$progId" }
    }

    # 5. 清理关联的ProgID（用户级）
    $userClassesPath = "HKCU:\Software\Classes\$extWithDot"
    if (Test-Path $userClassesPath) {
        $userProgId = (Get-ItemProperty $userClassesPath -ErrorAction SilentlyContinue).'(Default)'
        if ($userProgId) { Remove-RegistryPath "HKCU:\Software\Classes\$userProgId" }
    }
}

# 重启资源管理器使更改生效
try {
    $restartMsg = "`n----- Restarting Explorer -----"
    $log += $restartMsg
    Write-Host $restartMsg  # 控制台显示重启提示

    Stop-Process -Name explorer -Force -ErrorAction Stop
    Start-Process explorer -ErrorAction Stop
    
    $restartSuccessMsg = "Explorer restarted successfully"
    $log += $restartSuccessMsg
    Write-Host $restartSuccessMsg  # 控制台显示重启成功
}
catch {
    $restartFailMsg = "Warning: Failed to restart Explorer. Please restart your PC manually."
    $log += $restartFailMsg
    Write-Host $restartFailMsg  # 控制台显示重启失败
}

# 完成并保存日志
$endMsg = "`n===== Cleanup completed: $(Get-Date) ====="
$log += $endMsg
Write-Host $endMsg  # 控制台显示结束信息

$logPath = "$env:USERPROFILE\Desktop\FileAssociationCleanupLog.txt"
$log | Out-File -FilePath $logPath -Encoding UTF8
Write-Host "`nOperation completed! Log saved to: $logPath`n"  # 控制台显示日志路径