# 共享本次 COM 实例的身份登记与独立清理步骤。
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class K12OfficeWindow {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
"@

function New-K12OfficeApplication([string]$Application, [string]$OwnershipDirectory) {
    $processName = @{word='WINWORD';powerpoint='POWERPNT';excel='EXCEL'}[$Application]
    if (Get-Process -Name $processName -ErrorAction SilentlyContinue) {
        throw '请先关闭相应的 Office 应用，再执行本地转换。'
    }
    $starting = Join-Path $OwnershipDirectory "$Application.starting"
    [IO.File]::WriteAllText($starting, $Application)
    $started = [DateTime]::UtcNow
    $progId = @{word='Word.Application';powerpoint='PowerPoint.Application';excel='Excel.Application'}[$Application]
    $instance = New-Object -ComObject $progId
    # 窗口句柄直接关联 COM 实例，避免按名称猜测进程归属。
    [uint32]$instanceId = 0
    [void][K12OfficeWindow]::GetWindowThreadProcessId([IntPtr]$instance.Hwnd, [ref]$instanceId)
    if ($instanceId -eq 0) { throw '无法确认 Office 实例归属，请人工检查 Office 后重试。' }
    $process = Get-Process -Id $instanceId -ErrorAction Stop
    if ($process.ProcessName -ne $processName -or $process.StartTime.ToUniversalTime() -lt $started) {
        throw 'Office 实例并非本次启动，未操作该实例，请人工检查。'
    }
    $record = @{application=$Application;pid=[int]$instanceId;creation_time=[string]$process.StartTime.ToUniversalTime().ToFileTimeUtc()}
    $path = Join-Path $OwnershipDirectory "$Application.json"
    $record | ConvertTo-Json -Compress | Set-Content -LiteralPath "$path.tmp" -Encoding UTF8
    Move-Item -LiteralPath "$path.tmp" -Destination $path
    Remove-Item -LiteralPath $starting
    return $instance
}

function Close-K12OfficeObjects($Document, $Application, [string]$Kind) {
    # 每一步独立保护；关闭文档失败也必须继续退出并释放应用。
    if ($null -ne $Document) {
        try {
            if ($Kind -eq 'powerpoint') { $Document.Close() } else { $Document.Close($false) }
        } catch { [Console]::Error.WriteLine('Office 文档关闭失败，将继续清理本次实例。') }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Document) }
        catch { [Console]::Error.WriteLine('Office 文档引用释放失败。') }
    }
    if ($null -ne $Application) {
        try { $Application.Quit() }
        catch { [Console]::Error.WriteLine('Office 退出失败，将由监管进程清理本次实例。') }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Application) }
        catch { [Console]::Error.WriteLine('Office 应用引用释放失败。') }
    }
}
