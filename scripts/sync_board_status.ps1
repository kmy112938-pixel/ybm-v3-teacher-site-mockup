# Y클라우드 게시판 현황판 수동 동기화 (로컬 테스트/디버깅용)
# 정기 자동 동기화는 .github/workflows/sync-board-status.yml (GitHub Actions, 30분마다)이 담당한다.
# 이 스크립트는 PC와 무관하게 결과를 바로 확인하고 싶을 때 수동으로 실행하는 용도로만 남겨둔다.

$ErrorActionPreference = "Stop"

$RepoRoot = "C:\Users\YBM\project"
$SheetId = "1aF5gnTNRObLexo_QvDfOWHE1YzSCNCVY"
$SourceUrl = "https://docs.google.com/spreadsheets/d/$SheetId/export?format=xlsx"
$TempCopy = "$env:TEMP\ybm_board_status_sync.xlsx"
$LogFile = Join-Path $RepoRoot "scripts\sync.log"

function Write-Log($msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $LogFile -Value $line -Encoding utf8
}

try {
    # 1. 구글 시트(공유 링크)를 xlsx로 내려받기
    Invoke-WebRequest -Uri $SourceUrl -OutFile $TempCopy -UseBasicParsing

    # 2. HTML 재생성
    Set-Location $RepoRoot
    $result = & python scripts\generate_board_status.py $TempCopy docs\board-status.html "Y클라우드_게시판 현황_원본데이터 (팀 공유 구글시트)"
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
