<#
.SYNOPSIS
    xlsm ファイルの VBAProject をエクスポート／インポートします。
.DESCRIPTION
    xlsm-manager.xlsm を使わず、対象ブックだけを Excel COM で開いて
    VBAProject のエクスポートまたはインポートを実行します。
.PARAMETER Command
    実行するコマンド。"export" または "import" を指定します。
.PARAMETER TargetPath
    対象の xlsm ファイルのパス。
.PARAMETER Help
    このスクリプトのヘルプを表示します。
.EXAMPLE
    .\xlsm-manager.ps1 export "C:\work\target.xlsm"
    .\xlsm-manager.ps1 import "C:\work\target.xlsm"
#>
param(
    [Parameter(Position=0)]
    [ValidateSet('export','import')]
    [string]$Command,

    [Parameter(Position=1)]
    [string]$TargetPath,

    [switch]$Help
)

# ヘルプ表示
if ($Help -or -not $Command) {
    Write-Host @"
xlsm-manager.ps1 - xlsm ファイルの VBAProject 管理ツール

使い方:
  .\xlsm-manager.ps1 <Command> <TargetPath> [Options]

コマンド:
  export  VBAProject をテキストファイルへエクスポートします
  import  テキストファイルから VBAProject へインポートします

引数:
  TargetPath  対象の .xlsm ファイルのパス（相対・絶対どちらも可）

オプション:
  -Help       このヘルプを表示します

例:
  .\xlsm-manager.ps1 export "C:\work\MyMacro.xlsm"
  .\xlsm-manager.ps1 import "C:\work\MyMacro.xlsm"
"@
    exit 0
}

if (-not $TargetPath) {
    Write-Error "TargetPath が指定されていません。"
    exit 1
}

# PowerShell の出力エンコーディングを UTF-8 に設定（WSL 等での文字化け対策）
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# ヘルパー関数
# ---------------------------------------------------------------------------

function Ensure-Dir([string]$path) {
    if (-not (Test-Path $path)) { New-Item -ItemType Directory -Path $path | Out-Null }
}

# .cls ファイルから VERSION/Attribute ヘッダを除いたコード本体を返す
function Read-ImportableCode([string]$path) {
    $lines    = [System.IO.File]::ReadAllLines($path, [System.Text.Encoding]::Default)
    $inHeader = $false
    $result   = [System.Collections.Generic.List[string]]::new()

    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ($trimmed -like 'VERSION *') { $inHeader = $true; continue }
        if ($inHeader) {
            if ($trimmed -eq 'END') { $inHeader = $false }
            continue
        }
        if ($trimmed -like 'Attribute VB_*') { continue }
        $result.Add($line)
    }

    while ($result.Count -gt 0 -and $result[$result.Count - 1].Trim() -eq '') {
        $result.RemoveAt($result.Count - 1)
    }
    return $result -join "`r`n"
}

function Find-VBComponent($vbComponents, [string]$name) {
    foreach ($c in $vbComponents) {
        if ([string]::Compare($c.Name, $name, $true) -eq 0) { return $c }
    }
    return $null
}

# ---------------------------------------------------------------------------
# パス解決
# ---------------------------------------------------------------------------
$TargetPath = [System.IO.Path]::GetFullPath($TargetPath)
if (-not (Test-Path $TargetPath)) {
    Write-Error "ファイルが存在しません: $TargetPath"
    exit 1
}

$workbookFileName = [System.IO.Path]::GetFileName($TargetPath)
$rootDir          = [System.IO.Path]::GetDirectoryName($TargetPath)
$vbaProjectDir    = Join-Path $rootDir "VBAProject($workbookFileName)"
$excelObjectsDir  = Join-Path $vbaProjectDir "Microsoft Excel Objects"
$stdModuleDir     = Join-Path $vbaProjectDir "標準モジュール"
$classModuleDir   = Join-Path $vbaProjectDir "クラス モジュール"

# ---------------------------------------------------------------------------
# Excel 起動
# ---------------------------------------------------------------------------
$excel = New-Object -ComObject Excel.Application
$excel.Visible       = $false
$excel.DisplayAlerts = $false

# ===========================================================================
# export
# ===========================================================================
function Invoke-Export {
    foreach ($dir in @($vbaProjectDir, $excelObjectsDir, $stdModuleDir, $classModuleDir)) {
        Ensure-Dir $dir
    }

    $wb = $excel.Workbooks.Open($TargetPath, $true) # ReadOnly
    try {
        foreach ($comp in $wb.VBProject.VBComponents) {
            switch ($comp.Type) {
                100 { $exportPath = Join-Path $excelObjectsDir "$($comp.Name).cls" }  # Document
                  1 { $exportPath = Join-Path $stdModuleDir    "$($comp.Name).bas" }  # StdModule
                  2 { $exportPath = Join-Path $classModuleDir  "$($comp.Name).cls" }  # ClassModule
                default { $exportPath = $null }
            }
            if ($null -ne $exportPath) {
                if (Test-Path $exportPath) { Remove-Item $exportPath -Force }
                $comp.Export($exportPath) | Out-Null
            }
        }
        $wb.Close($false)
        Write-Host "export completed: $TargetPath"
    }
    catch {
        try { $wb.Close($false) } catch {}
        throw
    }
}

# ===========================================================================
# import
# ===========================================================================
function Invoke-Import {
    if (-not (Test-Path $vbaProjectDir)) {
        throw "インポート元フォルダが存在しません: $vbaProjectDir"
    }

    $timestamp   = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $backupParent = Join-Path $vbaProjectDir "._ImportBackup_"
    $backupRoot   = Join-Path $backupParent $timestamp
    $backupExcel  = Join-Path $backupRoot "Microsoft Excel Objects"
    $backupStd    = Join-Path $backupRoot "標準モジュール"
    $backupClass  = Join-Path $backupRoot "クラス モジュール"

    foreach ($dir in @($backupParent, $backupRoot, $backupExcel, $backupStd, $backupClass)) {
        Ensure-Dir $dir
    }

    $wb = $excel.Workbooks.Open($TargetPath, $false) # 書き込み可
    try {
        # --- 標準モジュール (.bas) ---
        if (Test-Path $stdModuleDir) {
            foreach ($file in Get-ChildItem $stdModuleDir -Filter "*.bas") {
                $modName  = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                $existing = Find-VBComponent $wb.VBProject.VBComponents $modName
                if ($null -ne $existing) {
                    $bak = Join-Path $backupStd "$($existing.Name).bas"
                    if (Test-Path $bak) { Remove-Item $bak -Force }
                    $existing.Export($bak) | Out-Null
                    $wb.VBProject.VBComponents.Remove($existing)
                }
                $imported = $wb.VBProject.VBComponents.Import($file.FullName)
                if ($imported.Type -ne 1) { throw "想定外のモジュール種別です: $($file.FullName)" }
            }
        }

        # --- クラスモジュール (.cls) ---
        if (Test-Path $classModuleDir) {
            foreach ($file in Get-ChildItem $classModuleDir -Filter "*.cls") {
                $modName  = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                $existing = Find-VBComponent $wb.VBProject.VBComponents $modName
                if ($null -ne $existing) {
                    $bak = Join-Path $backupClass "$($existing.Name).cls"
                    if (Test-Path $bak) { Remove-Item $bak -Force }
                    $existing.Export($bak) | Out-Null
                    $wb.VBProject.VBComponents.Remove($existing)
                }
                $imported = $wb.VBProject.VBComponents.Import($file.FullName)
                if ($imported.Type -ne 2) { throw "想定外のモジュール種別です: $($file.FullName)" }
            }
        }

        # --- Excel Objects (Document .cls) — コード本体のみ置換 ---
        if (Test-Path $excelObjectsDir) {
            foreach ($file in Get-ChildItem $excelObjectsDir -Filter "*.cls") {
                $modName  = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                $existing = Find-VBComponent $wb.VBProject.VBComponents $modName
                if ($null -eq $existing) { throw "対象の Excel Object が見つかりません: $modName" }
                if ($existing.Type -ne 100) { throw "対象が Excel Object ではありません: $modName" }

                $bak = Join-Path $backupExcel "$($existing.Name).cls"
                if (Test-Path $bak) { Remove-Item $bak -Force }
                $existing.Export($bak) | Out-Null

                $codeText = Read-ImportableCode $file.FullName
                $cm = $existing.CodeModule
                if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
                if ($codeText.Length -gt 0) { $cm.AddFromString($codeText) }
            }
        }

        $wb.Close($true) # 保存して閉じる
        Write-Host "import completed: $TargetPath"
    }
    catch {
        try { $wb.Close($false) } catch {}
        throw
    }
}

# ---------------------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------------------
try {
    switch ($Command) {
        'export' { Invoke-Export }
        'import' { Invoke-Import }
    }
    exit 0
}
catch {
    Write-Error "${Command} 中にエラーが発生しました: $_"
    exit 2
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
