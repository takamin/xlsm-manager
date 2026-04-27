xlsm-manager - マクロ付きエクセルファイルの管理ツール
=====================================================

概要
----

`xlsm-manager.ps1` は、Excel マクロ付きブック（`.xlsm`）の VBAProject をファイルシステムへエクスポート、またはファイルシステムからインポートする PowerShell スクリプトです。  
VBA コードをテキストファイルとして管理することで、Git 等のバージョン管理システムによる差分管理が可能になります。

---

前提条件
----

- Windows PowerShell 5.1 以上
- Microsoft Excel がインストールされていること（COM オートメーションを使用）
- 対象の `.xlsm` ファイルの VBAProject のアクセスが許可されていること  
  （Excel → ファイル → オプション → トラストセンター → マクロの設定 →「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効化）

---

インストール
----

インストールすることで、任意のディレクトリから `xlsm-manager` を呼び出せるようになります。

### PowerShell でインストールする

リポジトリのルートで以下を実行します。

```powershell
.\install.ps1
```

インストール先は `%USERPROFILE%\.local\bin` です。ユーザー環境変数 `PATH` に自動追加されます。  
新しい PowerShell ウィンドウを開くと以下のように使えます。

```powershell
xlsm-manager.ps1 export .\MyMacro.xlsm
xlsm-manager.ps1 import .\MyMacro.xlsm
```

**アンインストール:**

```powershell
.\install.ps1 -Uninstall
```

---

### Git Bash / WSL (Ubuntu) でインストールする

```bash
chmod +x install.sh
./install.sh
```

インストール先は `~/.local/bin` です。`~/.bashrc` または `~/.bash_profile` に `PATH` の設定が追記されます。  
新しいシェルを開くと以下のように使えます。

```bash
xlsm-manager export ./MyMacro.xlsm
xlsm-manager import ./MyMacro.xlsm
```

**アンインストール:**

```bash
./install.sh --uninstall
```

---

使い方（インストールなし）
----

```powershell
.\xlsm-manager.ps1 <Command> <TargetPath>
```

### パラメーター

| パラメーター | 必須 | 説明 |
| --- | --- | --- |
| `Command` | ○ | 実行するコマンド。`export` または `import` を指定します。 |
| `TargetPath` | ○ | 対象の `.xlsm` ファイルのパス（相対パス・絶対パスどちらも可）。 |

---

### Git Bash / WSL (Ubuntu) から使用する場合

`xlsm-manager.sh` を使用します。このスクリプトは実行時にプラットフォーム（Git Bash / WSL）を自動判定し、パス変換を行ったうえで `xlsm-manager.ps1` を呼び出します。

```bash
# export
./xlsm-manager.sh export ./target.xlsm

# import
./xlsm-manager.sh import ./target.xlsm
```

初回実行前に実行権限を付与してください。

```bash
chmod +x xlsm-manager.sh
```

#### パスの指定

Unix 形式のパス（相対・絶対どちらも可）をそのまま渡せます。Windows パスへの変換はスクリプトが自動的に行います。

| 環境 | パス例 |
| --- | --- |
| Git Bash | `./target.xlsm` または `/c/work/target.xlsm` |
| WSL (Ubuntu) | `./target.xlsm` または `/mnt/c/work/target.xlsm` |

#### プラットフォーム判定の仕組み

| 環境 | 判定方法 | パス変換コマンド |
| --- | --- | --- |
| Git Bash | `uname -s` が `MINGW*` / `MSYS*` / `CYGWIN*` | `cygpath -w` |
| WSL (Ubuntu) | `uname -s` が `Linux` かつ `/proc/version` に `microsoft` を含む | `wslpath -w` |

> **注意：** `powershell.exe` が PATH に存在しない場合は `/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe` を自動的に試みます。それでも見つからない場合はエラーメッセージが表示されます。

---

コマンド詳細
----

### export

VBAProject のすべてのモジュールをテキストファイルとしてエクスポートします。

```powershell
.\xlsm-manager.ps1 export "C:\work\target.xlsm"
```

**出力先フォルダ構造：**

```text
target.xlsm と同じフォルダ/
└── VBAProject(target.xlsm)/
    ├── Microsoft Excel Objects/   # シートやブックモジュール (.cls)
    ├── 標準モジュール/              # 標準モジュール (.bas)
    └── クラス モジュール/           # クラスモジュール (.cls)
```

- 対象ブックは読み取り専用で開かれ、エクスポート後に閉じられます。
- 同名のファイルが存在する場合は上書きされます。

---

### import

エクスポートしたテキストファイルを読み込み、VBAProject へ書き込みます。

```powershell
.\xlsm-manager.ps1 import "C:\work\target.xlsm"
```

**処理の流れ：**

1. `VBAProject(target.xlsm)` フォルダが存在することを確認します（存在しない場合はエラー）。
2. インポート前に既存のモジュールをタイムスタンプ付きバックアップフォルダへ退避します。
3. 各モジュールを置き換えます。
   - 標準モジュール（`.bas`）: 既存モジュールを削除してインポート
   - クラスモジュール（`.cls`）: 既存モジュールを削除してインポート
   - Excel Objects（`.cls`）: コード本体のみを置換（シート・ブックオブジェクト自体は削除不可のため）
4. ブックを上書き保存して閉じます。

**バックアップ先：**

```text
VBAProject(target.xlsm)/
└── ._ImportBackup_/
    └── yyyyMMdd_HHmmss/        # インポート実行時刻のタイムスタンプ
        ├── Microsoft Excel Objects/
        ├── 標準モジュール/
        └── クラス モジュール/
```

---

終了コード
----

| コード | 意味 |
| --- | --- |
| `0` | 正常終了 |
| `1` | 対象ファイルが存在しない |
| `2` | 処理中にエラーが発生した |

---

使用例
----

### PowerShell の場合

```powershell
# VBA をエクスポートしてバージョン管理に追加する
.\xlsm-manager.ps1 export "C:\work\MyMacro.xlsm"
git add "VBAProject(MyMacro.xlsm)"
git commit -m "update VBA"

# バージョン管理から取得したコードをブックへ反映する
git pull
.\xlsm-manager.ps1 import "C:\work\MyMacro.xlsm"
```

### Git Bash / WSL (Ubuntu) の場合

```bash
# VBA をエクスポートしてバージョン管理に追加する
./xlsm-manager.sh export ./MyMacro.xlsm
git add "VBAProject(MyMacro.xlsm)"
git commit -m "update VBA"

# バージョン管理から取得したコードをブックへ反映する
git pull
./xlsm-manager.sh import ./MyMacro.xlsm
```
