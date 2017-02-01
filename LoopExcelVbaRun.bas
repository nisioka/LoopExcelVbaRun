Attribute VB_Name = "Module1"
Option Explicit

Dim fu As FileUtil
Dim file As Variant
Dim formatFile As String
Dim formatDirectory As String
Dim workDirectory As String
Dim fromDirectory As String
Dim toDirectory As String

Private Sub initialize()
'
' 変数定義。
'
    Set fu = New FileUtil
    formatFile = "format.xlsm"
    formatDirectory = ThisWorkbook.Path & "\format\"
    workDirectory = ThisWorkbook.Path & "\work\"
    fromDirectory = ThisWorkbook.Path & "\..\1-fromExcel"
    toDirectory = ThisWorkbook.Path & "\..\2-toExcel\"
End Sub

Private Sub clean()
'
' ディレクトリ内ファイル削除。
'
    On Error Resume Next
    Kill workDirectory & "*"
    On Error Resume Next
    Kill toDirectory & "*"
End Sub

Public Sub BulkConversion()
Attribute BulkConversion.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BulkConversion Macro
'

    ' 画面描画しない
    Application.ScreenUpdating = False
    ' 確認メッセージ無効
    Application.DisplayAlerts = False
    
    Call initialize
    Call clean
    
    Dim result As Collection
    ' 対象ファイル一覧がフルパスで取得される
    Set result = fu.getFileListRecursive(fromDirectory, ".xls").Files

    For Each file In result
        
        ' フォーマットファイルをworkディレクトリにコピー
        FileCopy formatDirectory & formatFile, workDirectory & formatFile
         
        ' toブック用変数
        Dim toBook As Workbook
        Set toBook = Workbooks.Open(workDirectory & formatFile)

        ' 修正対象ファイルからコピー
        With Workbooks.Open(file)
            .Sheets("シート名").Copy Before:=toBook.Sheets(1)
            .Close
        End With
        ' フォーマットファイルのVBAを実行して再フォーマット
        Application.Run formatFile & "!ＶＢＡ名称"
        toBook.Close savechanges:=True
        
        ' ファイルの移動&リネーム
        Name workDirectory & formatFile As toDirectory & dir(file)
    Next

    ' 確認メッセージ有効
    Application.DisplayAlerts = True
    ' 画面描画する
    Application.ScreenUpdating = True
End Sub

