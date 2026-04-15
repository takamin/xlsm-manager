Attribute VB_Name = "modXLSM"
Option Explicit
' マクロ付きExcelファイルのVBAプロジェクトを管理するモジュール

' マクロ付きエクセルファイルのVBAプロジェクトをエクスポートするサブルーチン
' ========================================================================
'
' 引数
' ----
' filename: String - エクスポートするマクロ付きエクセルファイルのファイル名
'
' 戻り値
' ------
' Boolean - エクスポートが成功した場合はTrue、失敗した場合はFalseを返す
'
' エクスポート処理中にエラーが発生した場合は、エクスポートが失敗したことと、
' エラーの内容をメッセージボックスで表示しFalseを返す。
' リトライなどは行わない。
'
' 仕様詳細
' --------
'
' ### エクスポート先ルートフォルダ
'
' 出力先は、filenameと同じディレクトリに、`VBAProject(<bodyname>.xlsm)`という
' 名前のフォルダを作成し、その中にVBAプロジェクトのソースファイルの種別に応じて
' サブディレクトリを作成し、エクスポートする。
'
' ### Excelワークブックとワークシートのクラスモジュール
'
' 出力先サブディレクトリ: `Microsoft Excel Objects`
' 出力ファイル名: オブジェクト名＋拡張子
' 拡張子: `.cls`
'
' ### 標準モジュール
'
' 出力先サブディレクトリ: `標準モジュール`
' 出力ファイル名: モジュール名＋拡張子
' 拡張子: `.bas`
'
' ### クラスモジュール
'
' 出力先サブディレクトリ: `クラス モジュール`
' 出力ファイル名: モジュール名＋拡張子
' 拡張子: `.cls`
'
Public Function ExportAll(filename As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim wb As Workbook
    Dim targetPath As String
    Dim rootDir As String
    Dim workbookFileName As String
    Dim exportRoot As String
    Dim excelObjectsDir As String
    Dim stdModuleDir As String
    Dim classModuleDir As String
    Dim vbComp As Object
    Dim exportPath As String
    Dim openedByThisFunction As Boolean
    Dim i As Long

    If Len(Trim$(filename)) = 0 Then
        Err.Raise vbObjectError + 1000, , "filename が空です。"
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    targetPath = filename
    If Not IsAbsoluteWindowsPath(targetPath) Then
        targetPath = fso.BuildPath(ThisWorkbook.path, targetPath)
    End If

    If Not fso.FileExists(targetPath) Then
        Err.Raise vbObjectError + 1001, , "ファイルが存在しません: " & targetPath
    End If

    workbookFileName = fso.GetFileName(targetPath)
    rootDir = fso.GetParentFolderName(targetPath)

    For i = 1 To Application.Workbooks.Count
        If StrComp(Application.Workbooks(i).FullName, targetPath, vbTextCompare) = 0 Then
            Set wb = Application.Workbooks(i)
            Exit For
        End If
    Next i

    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(filename:=targetPath, ReadOnly:=True)
        openedByThisFunction = True
    End If

    exportRoot = fso.BuildPath(rootDir, "VBAProject(" & workbookFileName & ")")
    excelObjectsDir = fso.BuildPath(exportRoot, "Microsoft Excel Objects")
    stdModuleDir = fso.BuildPath(exportRoot, "標準モジュール")
    classModuleDir = fso.BuildPath(exportRoot, "クラス モジュール")

    If Not fso.FolderExists(exportRoot) Then fso.CreateFolder exportRoot
    If Not fso.FolderExists(excelObjectsDir) Then fso.CreateFolder excelObjectsDir
    If Not fso.FolderExists(stdModuleDir) Then fso.CreateFolder stdModuleDir
    If Not fso.FolderExists(classModuleDir) Then fso.CreateFolder classModuleDir

    For Each vbComp In wb.VBProject.VBComponents
        Select Case vbComp.Type
            Case 100 ' vbext_ct_Document
                exportPath = fso.BuildPath(excelObjectsDir, vbComp.Name & ".cls")
            Case 1 ' vbext_ct_StdModule
                exportPath = fso.BuildPath(stdModuleDir, vbComp.Name & ".bas")
            Case 2 ' vbext_ct_ClassModule
                exportPath = fso.BuildPath(classModuleDir, vbComp.Name & ".cls")
            Case Else
                exportPath = ""
        End Select

        If Len(exportPath) > 0 Then
            If fso.FileExists(exportPath) Then
                fso.DeleteFile exportPath, True
            End If
            vbComp.Export exportPath
        End If
    Next vbComp

    If openedByThisFunction Then
        wb.Close SaveChanges:=False
    End If

    ' エクスポートが成功した場合はTrueを返す
    ExportAll = True
    Exit Function

ErrorHandler:
    On Error Resume Next
    If openedByThisFunction Then
        wb.Close SaveChanges:=False
    End If
    On Error GoTo 0
    MsgBox "VBAProjectのエクスポート中にエラーが発生しました: " & Err.Description, vbExclamation
    ExportAll = False
End Function

Private Function IsAbsoluteWindowsPath(ByVal path As String) As Boolean
    Dim p As String
    Dim firstChar As String

    p = Trim$(path)
    If Len(p) = 0 Then
        IsAbsoluteWindowsPath = False
        Exit Function
    End If

    ' UNC パス: \\server\share\...
    If Left$(p, 2) = "\\" Then
        IsAbsoluteWindowsPath = True
        Exit Function
    End If

    ' ルート相対（\folder\... や /folder/...）
    If Left$(p, 1) = "\" Or Left$(p, 1) = "/" Then
        IsAbsoluteWindowsPath = True
        Exit Function
    End If

    ' ドライブ絶対パス: C:\folder\... または C:/folder/...
    If Len(p) >= 3 Then
        firstChar = Left$(p, 1)
        If Mid$(p, 2, 1) = ":" _
           And (Mid$(p, 3, 1) = "\" Or Mid$(p, 3, 1) = "/") _
           And firstChar Like "[A-Za-z]" Then
            IsAbsoluteWindowsPath = True
            Exit Function
        End If
    End If

    IsAbsoluteWindowsPath = False
End Function

Public Function ImportAll(filename As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim wb As Workbook
    Dim targetPath As String
    Dim rootDir As String
    Dim workbookFileName As String
    Dim importRoot As String
    Dim excelObjectsDir As String
    Dim stdModuleDir As String
    Dim classModuleDir As String
    Dim backupParent As String
    Dim backupRoot As String
    Dim backupExcelObjectsDir As String
    Dim backupStdModuleDir As String
    Dim backupClassModuleDir As String
    Dim openedByThisFunction As Boolean
    Dim i As Long

    If Len(Trim$(filename)) = 0 Then
        Err.Raise vbObjectError + 1100, , "filename が空です。"
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    targetPath = filename
    If Not IsAbsoluteWindowsPath(targetPath) Then
        targetPath = fso.BuildPath(ThisWorkbook.path, targetPath)
    End If

    If Not fso.FileExists(targetPath) Then
        Err.Raise vbObjectError + 1101, , "ファイルが存在しません: " & targetPath
    End If

    workbookFileName = fso.GetFileName(targetPath)
    rootDir = fso.GetParentFolderName(targetPath)
    importRoot = fso.BuildPath(rootDir, "VBAProject(" & workbookFileName & ")")
    excelObjectsDir = fso.BuildPath(importRoot, "Microsoft Excel Objects")
    stdModuleDir = fso.BuildPath(importRoot, "標準モジュール")
    classModuleDir = fso.BuildPath(importRoot, "クラス モジュール")

    If Not fso.FolderExists(importRoot) Then
        Err.Raise vbObjectError + 1102, , "インポート元フォルダが存在しません: " & importRoot
    End If

    For i = 1 To Application.Workbooks.Count
        If StrComp(Application.Workbooks(i).FullName, targetPath, vbTextCompare) = 0 Then
            Set wb = Application.Workbooks(i)
            Exit For
        End If
    Next i

    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(filename:=targetPath, ReadOnly:=False)
        openedByThisFunction = True
    ElseIf wb.ReadOnly Then
        Err.Raise vbObjectError + 1103, , "読み取り専用のためインポートできません: " & targetPath
    End If

    backupParent = fso.BuildPath(importRoot, "._ImportBackup_")
    backupRoot = fso.BuildPath(backupParent, Format$(Now, "yyyymmdd_hhnnss"))
    backupExcelObjectsDir = fso.BuildPath(backupRoot, "Microsoft Excel Objects")
    backupStdModuleDir = fso.BuildPath(backupRoot, "標準モジュール")
    backupClassModuleDir = fso.BuildPath(backupRoot, "クラス モジュール")

    EnsureFolderExists fso, backupParent
    EnsureFolderExists fso, backupRoot
    EnsureFolderExists fso, backupExcelObjectsDir
    EnsureFolderExists fso, backupStdModuleDir
    EnsureFolderExists fso, backupClassModuleDir

    ImportStandardModules wb, fso, stdModuleDir, backupStdModuleDir
    ImportClassModules wb, fso, classModuleDir, backupClassModuleDir
    ImportExcelObjects wb, fso, excelObjectsDir, backupExcelObjectsDir

    If openedByThisFunction Then
        wb.Close SaveChanges:=True
    End If

    ImportAll = True
    Exit Function

ErrorHandler:
    On Error Resume Next
    If openedByThisFunction Then
        wb.Close SaveChanges:=False
    End If
    On Error GoTo 0
    MsgBox "VBAProjectのインポート中にエラーが発生しました: " & Err.Description, vbExclamation
    ImportAll = False
End Function

Private Sub EnsureFolderExists(ByVal fso As Object, ByVal folderPath As String)
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

Private Sub ImportStandardModules(ByVal wb As Workbook, ByVal fso As Object, ByVal sourceDir As String, ByVal backupDir As String)
    ImportModulesFromFolder wb, fso, sourceDir, backupDir, ".bas", 1
End Sub

Private Sub ImportClassModules(ByVal wb As Workbook, ByVal fso As Object, ByVal sourceDir As String, ByVal backupDir As String)
    ImportModulesFromFolder wb, fso, sourceDir, backupDir, ".cls", 2
End Sub

Private Sub ImportModulesFromFolder(ByVal wb As Workbook, ByVal fso As Object, ByVal sourceDir As String, ByVal backupDir As String, ByVal ext As String, ByVal expectedType As Long)
    Dim fileObj As Object
    Dim moduleName As String
    Dim sourcePath As String
    Dim backupPath As String
    Dim existingComp As Object

    If Not fso.FolderExists(sourceDir) Then
        Exit Sub
    End If

    For Each fileObj In fso.GetFolder(sourceDir).Files
        If LCase$(fso.GetExtensionName(fileObj.Name)) = Mid$(ext, 2) Then
            sourcePath = fileObj.path
            moduleName = fso.GetBaseName(fileObj.Name)

            Set existingComp = FindVBComponentByName(wb, moduleName)
            If Not existingComp Is Nothing Then
                backupPath = fso.BuildPath(backupDir, existingComp.Name & ext)
                If fso.FileExists(backupPath) Then
                    fso.DeleteFile backupPath, True
                End If
                existingComp.Export backupPath
                wb.VBProject.VBComponents.Remove existingComp
            End If

            Set existingComp = wb.VBProject.VBComponents.Import(sourcePath)
            If existingComp.Type <> expectedType Then
                Err.Raise vbObjectError + 1110, , "想定外のモジュール種別です: " & sourcePath
            End If
        End If
    Next fileObj
End Sub

Private Sub ImportExcelObjects(ByVal wb As Workbook, ByVal fso As Object, ByVal sourceDir As String, ByVal backupDir As String)
    Dim fileObj As Object
    Dim moduleName As String
    Dim sourcePath As String
    Dim backupPath As String
    Dim existingComp As Object
    Dim codeText As String

    If Not fso.FolderExists(sourceDir) Then
        Exit Sub
    End If

    For Each fileObj In fso.GetFolder(sourceDir).Files
        If LCase$(fso.GetExtensionName(fileObj.Name)) = "cls" Then
            sourcePath = fileObj.path
            moduleName = fso.GetBaseName(fileObj.Name)

            Set existingComp = FindVBComponentByName(wb, moduleName)
            If existingComp Is Nothing Then
                Err.Raise vbObjectError + 1111, , "対象の Excel Object が見つかりません: " & moduleName
            End If

            If existingComp.Type <> 100 Then
                Err.Raise vbObjectError + 1112, , "対象が Excel Object ではありません: " & moduleName
            End If

            backupPath = fso.BuildPath(backupDir, existingComp.Name & ".cls")
            If fso.FileExists(backupPath) Then
                fso.DeleteFile backupPath, True
            End If
            existingComp.Export backupPath

            codeText = ReadImportableCodeFromClassFile(fso, sourcePath)
            ReplaceCodeModule existingComp.codeModule, codeText
        End If
    Next fileObj
End Sub

Private Function FindVBComponentByName(ByVal wb As Workbook, ByVal componentName As String) As Object
    Dim vbComp As Object

    For Each vbComp In wb.VBProject.VBComponents
        If StrComp(vbComp.Name, componentName, vbTextCompare) = 0 Then
            Set FindVBComponentByName = vbComp
            Exit Function
        End If
    Next vbComp

    Set FindVBComponentByName = Nothing
End Function

Private Sub ReplaceCodeModule(ByVal codeModule As Object, ByVal codeText As String)
    With codeModule
        If .CountOfLines > 0 Then
            .DeleteLines 1, .CountOfLines
        End If
        If Len(codeText) > 0 Then
            .AddFromString codeText
        End If
    End With
End Sub

Private Function ReadImportableCodeFromClassFile(ByVal fso As Object, ByVal sourcePath As String) As String
    Dim ts As Object
    Dim allText As String
    Dim lines() As String
    Dim i As Long
    Dim t As String
    Dim inHeaderBlock As Boolean
    Dim result As String

    Set ts = fso.OpenTextFile(sourcePath, 1, False)
    allText = ts.ReadAll
    ts.Close

    lines = Split(allText, vbCrLf)

    For i = LBound(lines) To UBound(lines)
        t = Trim$(lines(i))

        If Left$(t, 7) = "VERSION" Then
            inHeaderBlock = True
            GoTo ContinueLoop
        End If

        If inHeaderBlock Then
            If t = "END" Then
                inHeaderBlock = False
            End If
            GoTo ContinueLoop
        End If

        If Left$(t, 12) = "Attribute VB_" Then
            GoTo ContinueLoop
        End If

        If Len(result) = 0 Then
            result = lines(i)
        Else
            result = result & vbCrLf & lines(i)
        End If

ContinueLoop:
    Next i

    ReadImportableCodeFromClassFile = result
End Function

' ファイルダイアログを用いて、任意のエクセルマクロファイルを指定し、VBAProjectをインポートするサブルーチン
' ========================================================================
'
' 引数
' ----
' なし
'
' 戻り値
' ------
' なし
'
' 動作概要
' -------
' 1. ファイルダイアログを開いて、ユーザにインポートするファイルを指定してもらう。
' 2. 実行前にメッセージボックスで実行確認する。
' 3. 実行結果をメッセージボックスで表示する。
'
Public Sub ExportProjects()
    On Error GoTo ErrorHandler
    
    Dim fileDialog As Object
    Dim selectedFile As String
    Dim confirmResult As VbMsgBoxResult
    Dim exportResult As Boolean
    Dim resultMessage As String
    
    ' ファイルダイアログを開く
    Set fileDialog = Application.fileDialog(3) ' 3 = msoFileDialogFilePicker
    
    With fileDialog
        .Title = "エクスポートするVBAプロジェクトを含むExcelファイルを選択"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel マクロ有効ファイル (*.xlsm;*.xlm)", "*.xlsm;*.xlm"
        .Filters.Add "すべてのファイル (*.*)", "*.*"
        .FilterIndex = 1
        .InitialFileName = ""
        .Show
        
        If .SelectedItems.Count = 0 Then
            ' ファイルが選択されなかった場合は終了
            Exit Sub
        End If
        
        selectedFile = .SelectedItems(1)
    End With
    
    ' 実行確認メッセージボックス
    confirmResult = MsgBox( _
        "次のファイルからVBAプロジェクトをエクスポートします。" & vbCrLf & vbCrLf & _
        selectedFile & vbCrLf & vbCrLf & _
        "よろしいですか？", _
        vbExclamation + vbYesNo, _
        "エクスポート確認")
    
    If confirmResult <> vbYes Then
        ' ユーザーが「いいえ」を選択した場合は終了
        MsgBox "エクスポートがキャンセルされました。", vbInformation
        Exit Sub
    End If
    
    ' VBAプロジェクトをエクスポート
    exportResult = ExportAll(selectedFile)
    
    ' 結果をメッセージボックスで表示
    If exportResult Then
        resultMessage = "VBAプロジェクトのエクスポートが正常に完了しました。" & vbCrLf & vbCrLf & _
                        "ファイル: " & selectedFile
        MsgBox resultMessage, vbInformation, "エクスポート完了"
    Else
        ' エラーメッセージはExportAll関数内で表示されるため、ここでは追加表示しない
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "VBAプロジェクトのエクスポートが失敗しました。" & vbCrLf & vbCrLf & _
            "ファイル: " & selectedFile & vbCrLf & vbCrLf & _
            "予期しないエラーが発生しました: " & Err.Description, vbExclamation, "エクスポート失敗"
End Sub

Public Sub ImportProjects()
    On Error GoTo ErrorHandler
    
    Dim fileDialog As Object
    Dim selectedFile As String
    Dim confirmResult As VbMsgBoxResult
    Dim importResult As Boolean
    Dim resultMessage As String
    
    ' ファイルダイアログを開く
    Set fileDialog = Application.fileDialog(3) ' 3 = msoFileDialogFilePicker
    
    With fileDialog
        .Title = "インポートするVBAプロジェクトを含むExcelファイルを選択"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel マクロ有効ファイル (*.xlsm;*.xlm)", "*.xlsm;*.xlm"
        .Filters.Add "すべてのファイル (*.*)", "*.*"
        .FilterIndex = 1
        .InitialFileName = ""
        .Show
        
        If .SelectedItems.Count = 0 Then
            ' ファイルが選択されなかった場合は終了
            Exit Sub
        End If
        
        selectedFile = .SelectedItems(1)
    End With
    
    ' 実行確認メッセージボックス
    confirmResult = MsgBox( _
        "次のファイルからVBAプロジェクトをインポートします。" & vbCrLf & vbCrLf & _
        selectedFile & vbCrLf & vbCrLf & _
        "よろしいですか？", _
        vbExclamation + vbYesNo, _
        "インポート確認")
    
    If confirmResult <> vbYes Then
        ' ユーザーが「いいえ」を選択した場合は終了
        MsgBox "インポートがキャンセルされました。", vbInformation
        Exit Sub
    End If
    
    ' VBAプロジェクトをインポート
    importResult = ImportAll(selectedFile)
    
    ' 結果をメッセージボックスで表示
    If importResult Then
        resultMessage = "VBAプロジェクトのインポートが正常に完了しました。" & vbCrLf & vbCrLf & _
                        "ファイル: " & selectedFile
        MsgBox resultMessage, vbInformation, "インポート完了"
    Else
        ' エラーメッセージはImportAll関数内で表示されるため、ここでは追加表示しない
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "VBAプロジェクトのインポートが失敗しました。" & vbCrLf & vbCrLf & _
            "ファイル: " & selectedFile & vbCrLf & vbCrLf & _
            "予期しないエラーが発生しました: " & Err.Description, vbExclamation, "インポート失敗"
End Sub
