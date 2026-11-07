# 프로젝트 폴더에서 이 파일을 실행하면 로컬 웹 서비스가 시작됩니다.
$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'webapp\run_local.ps1')
