param(
    [Parameter(Mandatory=$true)][ValidateSet('word','powerpoint','excel')][string]$Application,
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [Parameter(Mandatory=$true)][string]$OwnershipDirectory,
    [Parameter(Mandatory=$true)][string]$TargetPath
)

# 路径作为独立参数传入，避免中文、空格或特殊字符成为脚本代码。
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'windows_office_lifecycle.ps1')
$officeApp = $null
$officeDocument = $null
try {
    switch ($Application) {
        'word' {
            $officeApp = New-K12OfficeApplication 'word' $OwnershipDirectory
            $officeApp.Visible = $false
            $officeApp.DisplayAlerts = 0
            $officeApp.AutomationSecurity = 3
            $officeDocument = $officeApp.Documents.Open($SourcePath, $false, $true, $false)
            $officeDocument.SaveAs2($TargetPath, 12)
        }
        'powerpoint' {
            $officeApp = New-K12OfficeApplication 'powerpoint' $OwnershipDirectory
            $officeApp.DisplayAlerts = 1
            $officeApp.AutomationSecurity = 3
            $officeDocument = $officeApp.Presentations.Open($SourcePath, -1, 0, 0)
            $officeDocument.SaveAs($TargetPath, 24)
        }
        'excel' {
            $officeApp = New-K12OfficeApplication 'excel' $OwnershipDirectory
            $officeApp.Visible = $false
            $officeApp.DisplayAlerts = $false
            $officeApp.AutomationSecurity = 3
            $officeApp.AskToUpdateLinks = $false
            $officeDocument = $officeApp.Workbooks.Open($SourcePath, 0, $true)
            $officeDocument.SaveAs($TargetPath, 51)
        }
    }
} catch {
    # 错误详情由 Python 层脱敏，不向网页返回本地源文件路径。
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    Close-K12OfficeObjects $officeDocument $officeApp $Application
}
