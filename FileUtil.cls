VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ファイル一覧を再帰的に取得する関数
' http://qiita.com/Tterima/items/72d56ee5e2be0a77c898

Option Explicit

Private m_fso As FileSystemObject
Private m_files As Collection

Property Get FSO() As FileSystemObject
    Set FSO = m_fso
End Property

Property Get Files() As Collection
    Set Files = m_files
End Property

' ファイル一覧を再帰的に取得する関数
' 引数： folder_path 取得する起点のフォルダ
' 引数： pattern 取得対象のパターン（正規表現）
Public Function getFileListRecursive(folder_path As String, Optional pattern As String = "") As FileUtil
    ' ファイル一覧の結果格納用
    Dim file_list_tmp As Collection

    ' ループ用変数の宣言
    Dim folder As Variant
    Dim file_path As Variant
    Dim dir As Variant

    ' オブジェクトの初期化
    Set file_list_tmp = New Collection

    ' 現在ディレクトリ内の全ファイルの取得
    For Each file_path In FSO.GetFolder(folder_path).Files
        If file_path Like "*" & pattern & "*" Then
            DoEvents    ' フリーズ防止用
            Call Files.Add(CStr(file_path))
        End If
    Next

    ' サブディレクトリの再帰
    For Each dir In FSO.GetFolder(folder_path).SubFolders
        Call getFileListRecursive(dir.Path, pattern)
    Next


    Set getFileListRecursive = Me
End Function

' メンバ変数の初期化
Private Sub Class_Initialize()
    Set m_fso = New FileSystemObject
    Set m_files = New Collection
End Sub

' メンバ変数の解放
Private Sub Class_Terminate()
    Set m_fso = Nothing
    Set m_files = Nothing
End Sub

