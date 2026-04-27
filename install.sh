#!/usr/bin/env bash
# install.sh
# xlsm-manager を Git Bash または WSL (Ubuntu) の環境にインストール、
# またはアンインストールします。
#
# 使い方:
#   ./install.sh [options]

set -euo pipefail

INSTALL_DIR="$HOME/.local/bin"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ACTION="help"  # デフォルトはヘルプ表示

# ---------------------------------------------------------------------------
# ヘルプ表示関数
# ---------------------------------------------------------------------------
show_help() {
    cat <<EOF
install.sh - xlsm-manager インストーラー

使い方:
  ./install.sh [options]

オプション:
  (引数なし)      このヘルプを表示します
  --install       xlsm-manager をインストールします
  --uninstall     xlsm-manager をアンインストールします
  --help, -h      このヘルプを表示します

説明:
  インストール先: ~/.local/bin
  - xlsm-manager.ps1
  - xlsm-manager.sh
  - xlsm-manager （拡張子なし版）

  ~/.bashrc または ~/.bash_profile に PATH の設定を追記します。

例:
  ./install.sh --install
  ./install.sh --uninstall
EOF
    exit 0
}

# ---------------------------------------------------------------------------
# 引数解析
# ---------------------------------------------------------------------------
if [[ $# -eq 0 ]]; then
    show_help
fi

for arg in "$@"; do
    case "$arg" in
        --install|-i)
            ACTION="install"
            ;;
        --uninstall|-u)
            ACTION="uninstall"
            ;;
        --help|-h)
            show_help
            ;;
        *)
            echo "Error: 不明なオプション: $arg" >&2
            exit 1
            ;;
    esac
done

# 複数の操作が指定されている場合はエラー
ACTION_COUNT=$(echo "$@" | grep -c -E "(--install|--uninstall)" || true)
if [[ $ACTION_COUNT -gt 1 ]]; then
    echo "Error: --install と --uninstall は同時に指定できません。" >&2
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
# PATH への追加・削除を ~/.bashrc / ~/.bash_profile へ書き込む関数
# ---------------------------------------------------------------------------
PATH_EXPORT_LINE='export PATH="$HOME/.local/bin:$PATH"'

rc_files() {
    # 優先順に返す
    local files=()
    [[ -f "$HOME/.bashrc" ]]       && files+=("$HOME/.bashrc")
    [[ -f "$HOME/.bash_profile" ]] && files+=("$HOME/.bash_profile")
    # どちらもなければ .bashrc を対象にする
    [[ ${#files[@]} -eq 0 ]]       && files+=("$HOME/.bashrc")
    echo "${files[@]}"
}

add_to_path_rc() {
    local added=false
    for rc in $(rc_files); do
        if ! grep -qF "$PATH_EXPORT_LINE" "$rc" 2>/dev/null; then
            printf '\n# xlsm-manager\n%s\n' "$PATH_EXPORT_LINE" >> "$rc"
            echo "PATH の設定を追加しました: $rc"
            added=true
        fi
    done
    if [[ "$added" == false ]]; then
        echo "PATH はすでに設定済みです。"
    fi
}

remove_from_path_rc() {
    for rc in $(rc_files); do
        if grep -qF "$PATH_EXPORT_LINE" "$rc" 2>/dev/null; then
            # xlsm-manager ブロック（コメント行 + export 行）を削除
            sed -i '/^# xlsm-manager$/d; /^'"$(echo "$PATH_EXPORT_LINE" | sed 's/[\/&]/\\&/g')"'$/d' "$rc"
            echo "PATH の設定を削除しました: $rc"
        fi
    done
}

# ---------------------------------------------------------------------------
# アンインストール
# ---------------------------------------------------------------------------
if [[ "$ACTION" == "uninstall" ]]; then
    echo "アンインストールを開始します..."

    for f in xlsm-manager.ps1 xlsm-manager.sh xlsm-manager; do
        dst="$INSTALL_DIR/$f"
        if [[ -f "$dst" ]]; then
            rm -f "$dst"
            echo "削除しました: $dst"
        fi
    done

    remove_from_path_rc

    echo ""
    echo "アンインストール完了。新しいシェルを開くと変更が反映されます。"
    exit 0
fi

# ---------------------------------------------------------------------------
# インストール実行のチェック
# ---------------------------------------------------------------------------
if [[ "$ACTION" != "install" ]]; then
    echo "Error: 無効なアクション: $ACTION" >&2
    exit 1
fi

# ---------------------------------------------------------------------------
# インストール
# ---------------------------------------------------------------------------
echo "インストールを開始します..."
echo "インストール先: $INSTALL_DIR"
echo ""

# インストール先ディレクトリの作成
mkdir -p "$INSTALL_DIR"

# xlsm-manager.ps1 のコピー
# Windows PowerShell 5.1 は BOM なし UTF-8 を CP932 として読み込むため、
# UTF-8 BOM (EF BB BF) を付与して保存する
PS1_SRC="$SCRIPT_DIR/xlsm-manager.ps1"
if [[ ! -f "$PS1_SRC" ]]; then
    echo "Error: インストール元が見つかりません: $PS1_SRC" >&2
    exit 1
fi

PS1_DST="$INSTALL_DIR/xlsm-manager.ps1"

python3 << PYEOF
import sys
try:
    with open('$PS1_SRC', 'r', encoding='utf-8-sig') as f:
        content = f.read()
    # utf-8-sig で書くと BOM 付き UTF-8 になる
    with open('$PS1_DST', 'w', encoding='utf-8-sig', newline='\r\n') as f:
        f.write(content)
except Exception as e:
    print(f"Error: {e}", file=sys.stderr)
    sys.exit(1)
PYEOF

echo "インストールしました: $PS1_DST"

# xlsm-manager.sh のコピー（拡張子ありと拡張子なしの両方）
SH_SRC="$SCRIPT_DIR/xlsm-manager.sh"
if [[ ! -f "$SH_SRC" ]]; then
    echo "Error: インストール元が見つかりません: $SH_SRC" >&2
    exit 1
fi
cp "$SH_SRC" "$INSTALL_DIR/xlsm-manager.sh"
cp "$SH_SRC" "$INSTALL_DIR/xlsm-manager"
chmod +x "$INSTALL_DIR/xlsm-manager.sh"
chmod +x "$INSTALL_DIR/xlsm-manager"
echo "インストールしました: $INSTALL_DIR/xlsm-manager"

# ---------------------------------------------------------------------------
# PATH への追加
# ---------------------------------------------------------------------------
add_to_path_rc

# 現在のセッションにも反映
export PATH="$INSTALL_DIR:$PATH"

# ---------------------------------------------------------------------------
# 完了メッセージ
# ---------------------------------------------------------------------------
echo ""
echo "インストール完了！"
echo ""
echo "使い方:"
echo "  xlsm-manager export <path/to/file.xlsm>"
echo "  xlsm-manager import <path/to/file.xlsm>"
echo ""
echo "※ 新しいシェルを開くと PATH が恒久的に有効になります。"
echo "  現在のシェルではすでに有効です。"
