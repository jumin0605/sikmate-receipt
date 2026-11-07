# 내 PC에서 웹 서버 켜기 (시험용)  ->  브라우저에서 http://127.0.0.1:8000
# 사용법:  PowerShell 에서  .\run_local.ps1
$ErrorActionPreference = 'Stop'
foreach ($e in 'PYTHONHOME','PYTHONNOUSERSITE','VIRTUAL_ENV','PIP_PREFIX','PIP_CACHE_DIR','OPENSSL_CONF','PYTHONPATH') {
    Remove-Item "Env:$e" -ErrorAction SilentlyContinue
}
# 프로젝트 .env는 app.py가 자동으로 읽는다. 다른 계정의 사용자 키를 덮어쓰지 않는다.
$projectEnv = Join-Path (Split-Path $PSScriptRoot -Parent) '.env'
$hasProjectKey = (Test-Path -LiteralPath $projectEnv) -and
    [bool](Select-String -LiteralPath $projectEnv -Pattern '^\s*(export\s+)?GEMINI_API_KEY\s*=\s*\S+' -Quiet)
if (-not $env:GEMINI_API_KEY -and -not $hasProjectKey) {
    $env:GEMINI_API_KEY = [Environment]::GetEnvironmentVariable('GEMINI_API_KEY', 'User')
}
if (-not $env:GEMINI_API_KEY -and -not $hasProjectKey) {
    throw '프로젝트 .env 또는 GEMINI_API_KEY 환경변수에 API 키를 설정해 주세요.'
}
if (-not $env:DATA_DIR) { $env:DATA_DIR = 'D:\_local_env\receipt_web_data' }
$env:PYTHONIOENCODING = 'utf-8'
# $env:APP_ACCESS_CODE = '여기에-접속코드'   # 켜면 코드를 아는 사람만 들어올 수 있음
Set-Location $PSScriptRoot
$projectPython = Join-Path (Split-Path $PSScriptRoot -Parent) '.venv\Scripts\python.exe'
$previousPython = 'D:\_local_env\receipt_web\Scripts\python.exe'
if (Test-Path -LiteralPath $projectPython) {
    $python = $projectPython
} elseif (Test-Path -LiteralPath $previousPython) {
    $python = $previousPython
} else {
    $python = (Get-Command python -ErrorAction Stop).Source
}
& $python -m uvicorn app:app --host 127.0.0.1 --port 8000
if ($LASTEXITCODE) { throw '웹 서버를 실행하지 못했습니다. 선택된 Python에 webapp/requirements.txt 패키지가 설치되어 있는지 확인하세요.' }
