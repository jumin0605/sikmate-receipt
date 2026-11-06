# 내 PC에서 웹 서버 켜기 (시험용)  ->  브라우저에서 http://127.0.0.1:8000
# 사용법:  PowerShell 에서  .\run_local.ps1
$ErrorActionPreference = 'Stop'
foreach ($e in 'PYTHONHOME','PYTHONNOUSERSITE','VIRTUAL_ENV','PIP_PREFIX','PIP_CACHE_DIR','OPENSSL_CONF','PYTHONPATH') {
    Remove-Item "Env:$e" -ErrorAction SilentlyContinue
}
if (-not $env:GEMINI_API_KEY) {
    $env:GEMINI_API_KEY = [Environment]::GetEnvironmentVariable('GEMINI_API_KEY', 'User')
}
if (-not $env:GEMINI_API_KEY) { throw 'GEMINI_API_KEY 환경변수가 없습니다.' }
if (-not $env:DATA_DIR) { $env:DATA_DIR = 'D:\_local_env\receipt_web_data' }
$env:PYTHONIOENCODING = 'utf-8'
# $env:APP_ACCESS_CODE = '여기에-접속코드'   # 켜면 코드를 아는 사람만 들어올 수 있음
Set-Location $PSScriptRoot
& 'D:\_local_env\receipt_web\Scripts\python.exe' -m uvicorn app:app --host 127.0.0.1 --port 8000
