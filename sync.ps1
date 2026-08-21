# 시연 폴더에서 최신 빌드를 가져와 이 배포 폴더에 반영한다.
# 실행: powershell -ExecutionPolicy Bypass -File sync.ps1  (이 폴더에서)
$ErrorActionPreference = 'Stop'
$src = Join-Path (Split-Path $PSScriptRoot -Parent) "시연"
Copy-Item (Join-Path $src "index.html") . -Force
Copy-Item (Join-Path $src "CRE_재무모델_빌더.html") . -Force
Copy-Item (Join-Path $src "업무_캘린더.html") . -Force
"복사 완료 — 이제 git add -A && git commit -m '...' && git push"
