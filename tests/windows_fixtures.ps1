param([Parameter(Mandatory=$true)][string]$FixtureDirectory, [Parameter(Mandatory=$true)][string]$OwnershipDirectory)
# 只在指定测试目录生成无宏的旧 Office 文档，不读取用户文档。
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../k12/windows_office_lifecycle.ps1')
$testWord = $null
$testDoc = $null
$testPowerPoint = $null
$testPresentation = $null
try {
    Write-Output 'Word: creating COM instance'
    $testWord = New-K12OfficeApplication 'word' $OwnershipDirectory
    Write-Output 'Word: creating document'
    $testWord.Visible = $false
    $testWord.DisplayAlerts = 0
    $testWord.AutomationSecurity = 3
    $testDoc = $testWord.Documents.Add()
    Write-Output 'Word: writing fixture'
    $testDoc.Content.Text = 'K12 Windows legacy Word integration test'
    $testDoc.SaveAs2((Join-Path $FixtureDirectory 'legacy.doc'), 0)
    Write-Output 'PowerPoint: creating COM instance'
    $testPowerPoint = New-K12OfficeApplication 'powerpoint' $OwnershipDirectory
    $testPowerPoint.AutomationSecurity = 3
    $testPresentation = $testPowerPoint.Presentations.Add(0)
    $testSlide = $testPresentation.Slides.Add(1, 1)
    $testSlide.Shapes.Title.TextFrame.TextRange.Text = 'K12 Windows legacy PPT integration test'
    $testPresentation.SaveAs((Join-Path $FixtureDirectory 'legacy.ppt'), 1)
    Write-Output 'Fixtures: ready'
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($testSlide)
} finally {
    Close-K12OfficeObjects $testDoc $testWord 'word'
    Close-K12OfficeObjects $testPresentation $testPowerPoint 'powerpoint'
}
