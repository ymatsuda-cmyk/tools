<#
.SYNOPSIS
  イヤホン切替(btswitch)の Windows 常駐エージェント。

.DESCRIPTION
  GAS に置かれた「いまどの端末が使うか」を一定間隔で見に行き、
  自分が指名されていれば接続、そうでなければ切断する。
  ブラウザから他の端末の Bluetooth は操作できないので、この形にしている。

  このファイルは必ず **UTF-8 (BOM付き)** で保存すること。
  Windows PowerShell 5.1 は BOM が無いと ANSI(CP932) として読むため、
  日本語の直後にある引用符や括弧を巻き込んで構文エラーになる。

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

.PARAMETER Service
  btcom に渡すプロファイルの短縮UUID。上から順に試し、最初に通ったものを以降使い続ける。
  110B=A2DP(音楽) / 1108=Headset / 111E=Handsfree。受け付ける組み合わせは機種によって違う。

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
  [string[]]$Service = @('110B', '1108', '111E'),
  [int]$IntervalSeconds = 10
)

$ErrorActionPreference = 'Stop'

# btcom は短縮UUIDを受け付けず、Bluetooth ベースUUID に埋めたフル GUID を要求する
$ServiceGuids = $Service | ForEach-Object { "{0000$($_.ToUpper())-0000-1000-8000-00805F9B34FB}" }

# 一度通ったプロファイルを覚えておく。毎回先頭から試すと切り替えがそのぶん遅くなる
$script:GoodGuid = $null

function Write-Log([string]$msg) {
  Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg"
}

function Invoke-Gas([hashtable]$body) {
  $body['token'] = $Token
  $json = $body | ConvertTo-Json -Compress
  # GASは正しいリクエストでもリダイレクト先が404を返すことがある。
  # ここで粘りすぎると切り替えが遅れるので、短く数回だけ。
  # それ以上続く不調は、10秒ごとに回る本体のループが拾い直す。
  $waits = @(1, 3)
  for ($i = 0; ; $i++) {
    # 毎回URLを変える。同じURLだと壊れたリダイレクトを掴んだまま繰り返す
    $sep = if ($GasUrl -like '*?*') { '&' } else { '?' }
    $url = $GasUrl + $sep + 'r=' + [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
    $res = $null
    try {
      # text/plain にしないと CORS ではなく GAS 側のリダイレクト処理で落ちることがある
      $res = Invoke-RestMethod -Uri $url -Method Post -ContentType 'text/plain;charset=utf-8' `
        -Body $json -MaximumRedirection 5
    } catch {
      if ($i -ge $waits.Count) { throw }
      Start-Sleep -Seconds $waits[$i]
      continue
    }
    # GAS が受け取ったうえで返したエラーは、投げ直しても同じなので即座に上げる
    if (-not $res.ok) { throw "GAS: $($res.error)" }
    return $res.data
  }
}

# ---- Bluetooth 操作 ----

function Get-BtDevices {
  # Bluetooth で見えている機器。名前の一部で絞る
  Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue |
    Where-Object { $_.FriendlyName -like "*$Name*" }
}

function Test-Connected {
  # 接続の有無は BTHENUM のサービスノードが居るかで見る。
  # SWD\MMDEVAPI のオーディオエンドポイントは切断後も残るため当てにならない
  $addr = ($Mac -replace '[^0-9A-Fa-f]', '').ToUpper()
  $nodes = Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue |
    Where-Object { $_.InstanceId -like 'BTHENUM\{*' -and $_.Status -eq 'OK' }
  if ($addr) {
    $nodes = $nodes | Where-Object { $_.InstanceId -match $addr }
  } else {
    $nodes = $nodes | Where-Object { $_.FriendlyName -like "*$Name*" }
  }
  # 110B=A2DP / 111E=Handsfree / 1108=Headset。どれか生きていれば使用中
  [bool]($nodes | Where-Object { $_.InstanceId -match '0000(110B|111E|1108)' })
}

function Use-Btcom { [bool](Get-Command btcom.exe -ErrorAction SilentlyContinue) }

function Invoke-Btcom([string]$flag, [switch]$All) {
  # 書式は btcom {-c|-r} -b<addr> -s<GUID>。成否は ERRORLEVEL で見る。
  # 標準エラーを 2>&1 で拾うと $ErrorActionPreference='Stop' の下では
  # 例外になって残りの候補を試せなくなるので、出力は捨てて終了コードだけ見る。
  $candidates = if ($script:GoodGuid) { @($script:GoodGuid) + $ServiceGuids } else { $ServiceGuids }
  $ok = $false
  foreach ($guid in ($candidates | Select-Object -Unique)) {
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
      & btcom.exe $flag "-b$Mac" "-s$guid" 2>$null | Out-Null
    } finally {
      $ErrorActionPreference = $prev
    }
    if ($LASTEXITCODE -eq 0) {
      $script:GoodGuid = $guid
      $ok = $true
      if (-not $All) { return $true }
    }
  }
  return $ok
}

# PnPデバイスの有効化 / 無効化。管理者権限が要るが、機種を選ばず効く
function Set-PnpConnected([bool]$on) {
  foreach ($d in Get-BtDevices) {
    if ($on) {
      Enable-PnpDevice -InstanceId $d.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
    } else {
      Disable-PnpDevice -InstanceId $d.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
    }
  }
}

function Connect-Earbuds {
  if ((Use-Btcom) -and $Mac -and (Invoke-Btcom '-c')) { return }
  Set-PnpConnected $true
}

function Disconnect-Earbuds {
  # 片方だけ落とすと Windows が張り直すので、音に関わるサービスは全部落とす
  if ((Use-Btcom) -and $Mac -and (Invoke-Btcom '-r' -All)) { return }
  # PnP方式は無効のままにしておく。有効に戻すと Windows がすぐ拾い直してしまい、
  # 他の端末に渡らない。自分の番が来たら Connect-Earbuds が有効に戻す
  Set-PnpConnected $false
}

# ---- 本体 ----

$mode = if ((Use-Btcom) -and $Mac) { 'btcom' } else { 'pnp(要管理者)' }
Write-Log "開始します: $DeviceId ($Label) / 操作方法: $mode / $IntervalSeconds 秒ごと"

$lastOwner = $null
try {
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
} finally {
  # 無効にしたまま終わると、次に手で有効化するまでイヤホンが使えなくなる
  if (-not ((Use-Btcom) -and $Mac)) {
    Write-Log 'デバイスを有効に戻します'
    Set-PnpConnected $true
  }
}
