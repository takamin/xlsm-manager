<#
.SYNOPSIS
    xlsm-manager をユーザー環境にインストール、またはアンインストールします。
.DESCRIPTION
    xlsm-manager.ps1 と xlsm-manager.sh を %USERPROFILE%\.local\bin へコピーし、
    ユーザー PATH に追加することで、任意のディレクトリから実行できるようにします。
.PARAMETER Uninstall
    インストール済みのファイルを削除し、PATH から除外します。
.PARAMETER Help
    このスクリプトのヘルプを表示します。
.EXAMPLE
    .\install.ps1
    .\install.ps1 -Uninstall
#>
param(
    [switch]$Uninstall,
    [switch]$Help
)

# ヘルプ表示
if ($Help) {
    Write-Host @"
install.ps1 - xlsm-manager インストーラー

使い方:
  .\install.ps1 [Options]

オプション:
  (引数なし)  xlsm-manager をユーザー環境にインストールします
  -Uninstall  インストール済みのファイルを削除します
  -Help       このヘルプを表示します

説明:
  インストール先: %USERPROFILE%\.local\bin
  - xlsm-manager.ps1
  - xlsm-manager.sh
  - xlsm-manager （拡張子なし版）

  ユーザー環境変数 PATH にインストール先ディレクトリを追加します。

例:
  .\install.ps1
  .\install.ps1 -Uninstall
"@
    exit 0
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$installDir = Join-Path $env:USERPROFILE ".local\bin"
$scriptDir  = $PSScriptRoot
$installedFiles = @(
    "xlsm-manager.ps1",
    "xlsm-manager.sh",
    "xlsm-manager"      # .sh の拡張子なしコピー（Git Bash / WSL 用）
)

# ---------------------------------------------------------------------------
# アンインストール
# ---------------------------------------------------------------------------
if ($Uninstall) {
    Write-Host "アンインストールを開始します..."

    foreach ($f in $installedFiles) {
        $dst = Join-Path $installDir $f
        if (Test-Path $dst) {
            Remove-Item $dst -Force
            Write-Host "削除しました: $dst"
        }
    }

    # PATH から除外
    $userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
    $entries  = $userPath -split ";" | Where-Object { $_ -ne $installDir -and $_ -ne "" }
    [Environment]::SetEnvironmentVariable("PATH", ($entries -join ";"), "User")
    Write-Host "PATH から除外しました: $installDir"

    Write-Host ""
    Write-Host "アンインストール完了。新しいシェルを開くと PATH の変更が反映されます。"
    exit 0
}

# ---------------------------------------------------------------------------
# インストール
# ---------------------------------------------------------------------------
Write-Host "インストールを開始します..."
Write-Host "インストール先: $installDir"
Write-Host ""

# インストール先ディレクトリの作成
if (-not (Test-Path $installDir)) {
    New-Item -ItemType Directory -Path $installDir | Out-Null
    Write-Host "ディレクトリを作成しました: $installDir"
}

# xlsm-manager.ps1 のコピー
$ps1Src = Join-Path $scriptDir "xlsm-manager.ps1"
$ps1Dst = Join-Path $installDir "xlsm-manager.ps1"
if (-not (Test-Path $ps1Src)) {
    Write-Error "インストール元が見つかりません: $ps1Src"
    exit 1
}
Copy-Item $ps1Src $ps1Dst -Force
Write-Host "インストールしました: $ps1Dst"

# xlsm-manager.sh のコピー（Git Bash / WSL 向け）
$shSrc = Join-Path $scriptDir "xlsm-manager.sh"
$shDst = Join-Path $installDir "xlsm-manager.sh"
if (Test-Path $shSrc) {
    Copy-Item $shSrc $shDst -Force
    Write-Host "インストールしました: $shDst"

    # 拡張子なしでも呼び出せるよう複製（Git Bash / WSL での xlsm-manager コマンド用）
    $noExtDst = Join-Path $installDir "xlsm-manager"
    Copy-Item $shSrc $noExtDst -Force
    Write-Host "インストールしました: $noExtDst"
} else {
    Write-Warning "xlsm-manager.sh が見つかりませんでした。Git Bash / WSL 用のインストールをスキップします。"
}

# ---------------------------------------------------------------------------
# ユーザー PATH への追加
# ---------------------------------------------------------------------------
$userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
$entries  = $userPath -split ";" | Where-Object { $_ -ne "" }

if ($entries -notcontains $installDir) {
    $newPath = ($entries + $installDir) -join ";"
    [Environment]::SetEnvironmentVariable("PATH", $newPath, "User")
    Write-Host ""
    Write-Host "PATH に追加しました: $installDir"
    Write-Host "※ 新しいシェルを開くと有効になります。"
    Write-Host "  現在のセッションで即座に有効にするには以下を実行してください:"
    Write-Host "  `$env:PATH += `";$installDir`""
} else {
    Write-Host ""
    Write-Host "PATH はすでに設定済みです: $installDir"
}

# ---------------------------------------------------------------------------
# 完了メッセージ
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "インストール完了！"
Write-Host ""
Write-Host "使い方（PowerShell）:"
Write-Host "  xlsm-manager.ps1 export <path\to\file.xlsm>"
Write-Host "  xlsm-manager.ps1 import <path\to\file.xlsm>"
Write-Host ""
Write-Host "使い方（Git Bash / WSL）:"
Write-Host "  xlsm-manager export <path/to/file.xlsm>"
Write-Host "  xlsm-manager import <path/to/file.xlsm>"
