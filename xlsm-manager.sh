#!/usr/bin/env bash
# xlsm-manager.sh
# Git Bash および WSL (Ubuntu) から xlsm-manager.ps1 を呼び出すラッパースクリプト。
# プラットフォームを自動判定し、パス変換を行ったうえで powershell.exe へ委譲します。
#
# 使い方:
#   ./xlsm-manager.sh export <path/to/target.xlsm>
#   ./xlsm-manager.sh import <path/to/target.xlsm>

set -euo pipefail

# シンボリックリンク経由で実行された場合も実体のディレクトリを取得する
_resolve_script_dir() {
    local src="$0"
    while [[ -L "$src" ]]; do
        local dir
        dir="$(cd "$(dirname "$src")" && pwd)"
        src="$(readlink "$src")"
        [[ "$src" != /* ]] && src="$dir/$src"
    done
    cd "$(dirname "$src")" && pwd
}

# ---------------------------------------------------------------------------
# ヘルプ表示
# ---------------------------------------------------------------------------
usage() {
    cat <<EOF
Usage: $(basename "$0") <command> <target.xlsm>

Commands:
  export  VBAProject をテキストファイルへエクスポートします
  import  テキストファイルから VBAProject へインポートします

Arguments:
  target.xlsm  対象の .xlsm ファイルのパス（相対・絶対どちらも可）

Examples:
  ./xlsm-manager.sh export ./MyMacro.xlsm
  ./xlsm-manager.sh import /mnt/c/work/MyMacro.xlsm
EOF
    exit 1
}

# ---------------------------------------------------------------------------
# 引数チェック
# ---------------------------------------------------------------------------
if [[ $# -eq 0 ]]; then
    usage
fi

# ヘルプオプションのチェック
for arg in "$@"; do
    case "$arg" in
        --help|-h)
            usage
            ;;
    esac
done

if [[ $# -ne 2 ]]; then
    usage
fi

COMMAND="$1"
TARGET_PATH="$2"

if [[ "$COMMAND" != "export" && "$COMMAND" != "import" ]]; then
    echo "Error: コマンドは 'export' または 'import' を指定してください。" >&2
    exit 1
fi

# ---------------------------------------------------------------------------
# プラットフォーム判定
# ---------------------------------------------------------------------------
detect_platform() {
    local kernel
    kernel="$(uname -s)"
    case "$kernel" in
        MINGW*|MSYS*|CYGWIN*)
            echo "gitbash"
            ;;
        Linux*)
            if grep -qi microsoft /proc/version 2>/dev/null; then
                echo "wsl"
            else
                echo "linux"
            fi
            ;;
        *)
            echo "unknown"
            ;;
    esac
}

PLATFORM="$(detect_platform)"

if [[ "$PLATFORM" == "linux" || "$PLATFORM" == "unknown" ]]; then
    echo "Error: このスクリプトは Git Bash または WSL (Ubuntu) 上での使用を前提としています。" >&2
    exit 1
fi

# ---------------------------------------------------------------------------
# Unix パス → Windows パスへ変換
# ---------------------------------------------------------------------------
to_win_path() {
    local path="$1"
    case "$PLATFORM" in
        gitbash)
            cygpath -w "$path"
            ;;
        wsl)
            wslpath -w "$path"
            ;;
    esac
}

# ---------------------------------------------------------------------------
# 絶対パスへ解決（realpath があれば使用、なければ手動解決）
# ---------------------------------------------------------------------------
resolve_path() {
    local path="$1"
    if command -v realpath &>/dev/null; then
        realpath "$path"
    else
        # Git Bash 向けフォールバック
        local dir file
        dir="$(cd "$(dirname "$path")" 2>/dev/null && pwd)"
        file="$(basename "$path")"
        echo "${dir}/${file}"
    fi
}

# ---------------------------------------------------------------------------
# powershell.exe の検索
# ---------------------------------------------------------------------------
find_powershell() {
    if command -v powershell.exe &>/dev/null; then
        echo "powershell.exe"
    elif [[ -f "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe" ]]; then
        echo "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
    else
        echo ""
    fi
}

POWERSHELL="$(find_powershell)"
if [[ -z "$POWERSHELL" ]]; then
    echo "Error: powershell.exe が見つかりません。" >&2
    echo "       PATH を確認するか、フルパスで指定してください:" >&2
    echo "       /mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe" >&2
    exit 1
fi

# ---------------------------------------------------------------------------
# パス変換
# ---------------------------------------------------------------------------
SCRIPT_DIR="$(_resolve_script_dir)"
SCRIPT_WIN="$(to_win_path "${SCRIPT_DIR}/xlsm-manager.ps1")"

# 対象ファイルの存在確認（Unix パスで行う）
TARGET_ABS="$(resolve_path "$TARGET_PATH")"
if [[ ! -f "$TARGET_ABS" ]]; then
    echo "Error: ファイルが存在しません: $TARGET_ABS" >&2
    exit 1
fi

TARGET_WIN="$(to_win_path "$TARGET_ABS")"

# ---------------------------------------------------------------------------
# 実行
# ---------------------------------------------------------------------------
exec "$POWERSHELL" -ExecutionPolicy Bypass -File "$SCRIPT_WIN" "$COMMAND" "$TARGET_WIN"
