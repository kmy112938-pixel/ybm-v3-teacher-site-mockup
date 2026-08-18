# Y클라우드 게시판 현황판 자동 동기화
# OneDrive 원본 엑셀을 읽어 docs/board-status.html을 재생성하고, 변경이 있으면 GitHub에 자동 push한다.
# Windows 작업 스케줄러가 30분마다 이 스크립트를 실행한다.

$ErrorActionPreference = "Stop"

$RepoRoot = "C:\Users\YBM\project"
$SourceXlsx = "C:\Users\YBM\OneDrive - YBM, Inc\01_실무\게시판운영 현황\Y클라우드_게시판 현황_원본데이터.xlsx"
$TempCopy = "$env:TEMP\ybm_board_status_sync.xlsx"
$LogFile = Join-Path $RepoRoot "scripts\sync.log"

function Write-Log($msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $LogFile -Value $line -Encoding utf8
}

try {
    # 1. OneDrive 클라우드 파일을 로컬로 강제 다운로드(hydrate)하며 복사
    Copy-Item -Path $SourceXlsx -Destination $TempCopy -Force

    # 2. HTML 재생성
    Set-Location $RepoRoot
    $result = & python scripts\generate_board_status.py $TempCopy
    if ($LASTEXITCODE -ne 0) {
        throw "generate_board_status.py 실패 (exit $LASTEXITCODE): $result"
    }
    Write-Log "generate result: $result"

    if ($result -notmatch "CHANGED") {
        Write-Log "변경 없음, 종료"
        exit 0
    }

    # 3. git commit & push (변경된 경우에만)
    git add docs\board-status.html
    $staged = git diff --cached --quiet; $hasChange = ($LASTEXITCODE -ne 0)

    if (-not $hasChange) {
        Write-Log "git 기준으로도 변경 없음, 종료"
        exit 0
    }

    $ts = Get-Date -Format "yyyy-MM-dd HH:mm"
    git commit -m "자동 동기화: 게시판 현황 갱신 ($ts)" | Out-Null
    git push origin master | Out-Null
    Write-Log "push 완료 ($ts)"
}
catch {
    Write-Log "오류: $($_.Exception.Message)"
    exit 1
}
