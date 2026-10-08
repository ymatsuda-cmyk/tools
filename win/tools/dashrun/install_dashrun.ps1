<#
  dashrun:// URL スキームを現在のユーザーに登録する（管理者権限は不要）

  使い方:
    .\install_dashrun.ps1                 登録
    .\install_dashrun.ps1 -Uninstall      登録解除
    .\install_dashrun.ps1 -PythonW "C:\Python312\pythonw.exe"   pythonw の場所を指定
#>
param(
  [switch]$Uninstall,
  [string]$PythonW
)
$ErrorActionPreference = 'Stop'
$key = 'HKCU:\Software\Classes\dashrun'

if ($Uninstall) {
  if (Test-Path $key) { Remove-Item -Path $key -Recurse -Force }
  Write-Host 'dashrun の登録を解除しました（許可済みリスト %APPDATA%\dashrun は残しています）'
  return
}

$handler = Join-Path $PSScriptRoot 'dashrun_handler.pyw'
if (-not (Test-Path $handler)) { throw "dashrun_handler.pyw が見つかりません: $handler" }

if (-not $PythonW) {
  foreach ($name in 'pythonw.exe', 'pyw.exe') {
    $cmd = Get-Command $name -ErrorAction SilentlyContinue
    if ($cmd) { $PythonW = $cmd.Source; break }
  }
}
if (-not $PythonW -or -not (Test-Path $PythonW)) {
  throw 'pythonw.exe が見つかりません。python.org 版の Python をインストールするか -PythonW で場所を指定してください'
}
if ($PythonW -like '*\WindowsApps\*') {
  Write-Warning 'Microsoft Store 版の Python が使われています。動かない場合は python.org 版をインストールし -PythonW で指定してください'
}

New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name '(default)' -Value 'URL:dashrun Protocol'
New-ItemProperty -Path $key -Name 'URL Protocol' -Value '' -PropertyType String -Force | Out-Null
New-Item -Path "$key\DefaultIcon" -Force | Out-Null
Set-ItemProperty -Path "$key\DefaultIcon" -Name '(default)' -Value "`"$PythonW`",0"
New-Item -Path "$key\shell\open\command" -Force | Out-Null
Set-ItemProperty -Path "$key\shell\open\command" -Name '(default)' -Value "`"$PythonW`" `"$handler`" `"%1`""

Write-Host 'dashrun を登録しました'
Write-Host "  pythonw : $PythonW"
Write-Host "  handler : $handler"
Write-Host ''
Write-Host '動作確認（メモ帳が起動すれば成功。初回は許可ダイアログが出ます）:'
Write-Host "  Start-Process 'dashrun://launch?path=C%3A%5CWindows%5Csystem32%5Cnotepad.exe'"
