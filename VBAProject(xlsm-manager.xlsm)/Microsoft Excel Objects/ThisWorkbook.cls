VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub ExportSystem()
    On Error GoTo ErrorHandler
    
    Dim confirmResult As VbMsgBoxResult
    Dim exportResult As Boolean
    Dim resultMessage As String
    
    ' 実行確認メッセージボックス
    confirmResult = MsgBox( _
        "このエクセルブック(xlsm-manager.xlsm)のVBAプロジェクトをエクスポートします。" & vbCrLf & vbCrLf & _
        "ファイル: xlsm-manager.xlsm" & vbCrLf & vbCrLf & _
        "よろしいですか？", _
        vbExclamation + vbYesNo, _
        "エクスポート確認")
    
    If confirmResult <> vbYes Then
        MsgBox "エクスポートがキャンセルされました。", vbInformation
        Exit Sub
    End If
    
    ' VBAプロジェクトをエクスポート
    exportResult = modXLSM.ExportAll("xlsm-manager.xlsm")
    
    ' 結果をメッセージボックスで表示
    If exportResult Then
        resultMessage = "VBAプロジェクトのエクスポートが正常に完了しました。" & vbCrLf & vbCrLf & _
                        "ファイル: xlsm-manager.xlsm"
        MsgBox resultMessage, vbInformation, "エクスポート完了"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "VBAプロジェクトのエクスポートが失敗しました。" & vbCrLf & vbCrLf & _
            "予期しないエラーが発生しました: " & Err.Description, vbExclamation, "エクスポート失敗"
End Sub

Sub ImportSystem()
    On Error GoTo ErrorHandler
    
    Dim confirmResult As VbMsgBoxResult
    Dim importResult As Boolean
    Dim resultMessage As String
    
    ' 実行確認メッセージボックス
    confirmResult = MsgBox( _
        "このエクセルブック(xlsm-manager.xlsm)のVBAプロジェクトをインポートします。" & vbCrLf & vbCrLf & _
        "ファイル: xlsm-manager.xlsm" & vbCrLf & vbCrLf & _
        "よろしいですか？", _
        vbExclamation + vbYesNo, _
        "インポート確認")
    
    If confirmResult <> vbYes Then
        MsgBox "インポートがキャンセルされました。", vbInformation
        Exit Sub
    End If
    
    ' VBAプロジェクトをインポート
    importResult = modXLSM.ImportAll("xlsm-manager.xlsm")
    
    ' 結果をメッセージボックスで表示
    If importResult Then
        resultMessage = "VBAプロジェクトのインポートが正常に完了しました。" & vbCrLf & vbCrLf & _
                        "ファイル: xlsm-manager.xlsm"
        MsgBox resultMessage, vbInformation, "インポート完了"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "VBAプロジェクトのインポートが失敗しました。" & vbCrLf & vbCrLf & _
            "予期しないエラーが発生しました: " & Err.Description, vbExclamation, "インポート失敗"
End Sub
