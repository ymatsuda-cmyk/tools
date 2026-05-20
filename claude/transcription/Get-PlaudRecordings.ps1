$OutputFile = "plaud_recordings.json"

Write-Host "▶ Plaud 録音一覧を取得します"
Write-Host "▶ 出力ファイル: $OutputFile"

if (-not $env:PLAUD_TOKEN) {
    Write-Error "PLAUD_TOKEN が未設定です"
    exit 1
}

if (-not $env:PLAUD_API_DOMAIN) {
    Write-Error "PLAUD_API_DOMAIN が未設定です"
    exit 1
}

plaud files list --all --json |
    Out-File -FilePath $OutputFile -Encoding utf8

if ($LASTEXITCODE -ne 0) {
    Write-Error "録音一覧の取得に失敗しました"
    exit 1
}

Write-Host "✅ 取得完了"
