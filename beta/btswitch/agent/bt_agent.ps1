<#
.SYNOPSIS
  イヤホン切替(btswitch)の Windows 常駐エージェント。

.DESCRIPTION
  GAS に置かれた「いまどの端末が使うか」を一定間隔で見に行き、
  自分が指名されていれば接続、そうでなければ切断する。
  ブラウザから他の端末の Bluetooth は操作できないので、この形にしている。

  接続 / 切断の手段は2つを自動で使い分ける:
    1. btcom.exe (Bluetooth Command Line Tools) があればそれを使う。管理者権限は不要
       https://bluetoothinstaller.com/bluetooth-command-line-tools
    2. 無ければ PnP デバイスの無効化 / 有効化で代用する。こちらは管理者権限が必要

.PARAMETER GasUrl
  GAS ウェブアプリのURL(/exec)

.PARAMETER Token
  GASの ACCESS_TOKEN と同じ合言葉

.PARAMETER DeviceId
  この端末のID。画面の設定と揃えると「この端末」と表示される(例: pc-home)

.PARAMETER Label
  画面に出す表示名(例: 自宅デスクトップ)

.PARAMETER Mac
  イヤホンのMACアドレス。btcom を使う場合に必要(例: 00:11:22:33:44:55)

.PARAMETER Name
  イヤホンのデバイス名の一部。PnP方式で使う(既定: Liberty)

.EXAMPLE
  .\bt_agent.ps1 -GasUrl https://script.google.com/macros/s/xxx/exec -Token himitsu `
                 -DeviceId pc-home -Label "自宅デスクトップ" -Mac 00:11:22:33:44:55
#>
param(
  [Parameter(Mandatory = $true)][string]$GasUrl,
  [Parameter(Mandatory = $true)][string]$Token,
  [Parameter(Mandatory = $true)][string]$DeviceId,
  [string]$Label = $env:COMPUTERNAME,
  [string]$Mac = '',
  [string]$Name = 'Liberty',
  [int]$IntervalSeconds = 10
)

$ErrorActionPreference = 'Stop'

function Write-Log([string]$msg) {
  Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg"
}

function Invoke-Gas([hashtable]$body) {
  $body['token'] = $Token
  # text/plain にしないと CORS ではなく GAS 側のリダイレクト処理で落ちることがある
  $res = Invoke-RestMethod -Uri $GasUrl -Method Post -ContentType 'text/plain;charset=utf-8' `
    -Body ($body | ConvertTo-Json -Compress) -MaximumRedirection 5
  if (-not $res.ok) { throw "GAS: $($res.error)" }
  return $res.data
}

# ---- Bluetooth 操作 ----

function Get-BtDevices {
  # Bluetooth で見えている機器。名前の一部で絞る
  Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue |
    Where-Object { $_.FriendlyName -like "*$Name*" }
}

function Test-Connected {
  $devices = Get-BtDevices
  if (-not $devices) { return $false }
  # 接続中はオーディオ側のエンドポイントが OK になる
  [bool]($devices | Where-Object { $_.Status -eq 'OK' -and $_.Class -in @('AudioEndpoint', 'MEDIA') })
}

function Use-Btcom { [bool](Get-Command btcom.exe -ErrorAction SilentlyContinue) }

function Connect-Earbuds {
  if ((Use-Btcom) -and $Mac) {
    # -s110b = A2DP(オーディオ)サービスに接続する
    & btcom.exe -b"$Mac" -s110b -c | Out-Null
    return
  }
  foreach ($d in Get-BtDevices) {
    if ($d.Status -ne 'OK') { Enable-PnpDevice -InstanceId $d.InstanceId -Confirm:$false -ErrorAction SilentlyContinue }
  }
}

function Disconnect-Earbuds {
  if ((Use-Btcom) -and $Mac) {
    & btcom.exe -b"$Mac" -s110b -d | Out-Null
    return
  }
  # PnP方式は「無効化して即有効化」で切断する。無効のままだと次に掴めない
  foreach ($d in Get-BtDevices) {
    Disable-PnpDevice -InstanceId $d.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
  }
  Start-Sleep -Seconds 2
  foreach ($d in Get-BtDevices) {
    Enable-PnpDevice -InstanceId $d.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
  }
}

# ---- 本体 ----

$mode = if ((Use-Btcom) -and $Mac) { 'btcom' } else { 'pnp(要管理者)' }
Write-Log "開始します: $DeviceId ($Label) / 操作方法: $mode / $IntervalSeconds 秒ごと"

$lastOwner = $null
while ($true) {
  try {
    $state = Invoke-Gas @{ action = 'getState' }
    $owner = [string]$state.owner
    $connected = Test-Connected

    if ($owner -eq $DeviceId -and -not $connected) {
      Write-Log '自分の番になりました。接続します'
      Connect-Earbuds
      Start-Sleep -Seconds 3
      $connected = Test-Connected
    } elseif ($owner -ne $DeviceId -and $connected) {
      Write-Log "他の端末($owner)に渡します。切断します"
      Disconnect-Earbuds
      Start-Sleep -Seconds 3
      $connected = Test-Connected
    }

    if ($owner -ne $lastOwner) {
      Write-Log "現在の使用端末: $(if ($owner) { $owner } else { '(なし)' })"
      $lastOwner = $owner
    }

    Invoke-Gas @{ action = 'heartbeat'; deviceId = $DeviceId; label = $Label; connected = $connected; note = $mode } | Out-Null
  } catch {
    Write-Log "!! $($_.Exception.Message)"
  }
  Start-Sleep -Seconds $IntervalSeconds
}
