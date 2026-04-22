# version.json 에서 version 필드를 읽어 stdout 으로 출력
$json = Get-Content -Raw -Path (Join-Path $PSScriptRoot "version.json") | ConvertFrom-Json
Write-Output $json.version
