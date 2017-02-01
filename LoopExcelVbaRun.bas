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
' �ϐ���`�B
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
' �f�B���N�g�����t�@�C���폜�B
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

    ' ��ʕ`�悵�Ȃ�
    Application.ScreenUpdating = False
    ' �m�F���b�Z�[�W����
    Application.DisplayAlerts = False
    
    Call initialize
    Call clean
    
    Dim result As Collection
    ' �Ώۃt�@�C���ꗗ���t���p�X�Ŏ擾�����
    Set result = fu.getFileListRecursive(fromDirectory, ".xls").Files

    For Each file In result
        
        ' �t�H�[�}�b�g�t�@�C����work�f�B���N�g���ɃR�s�[
        FileCopy formatDirectory & formatFile, workDirectory & formatFile
         
        ' to�u�b�N�p�ϐ�
        Dim toBook As Workbook
        Set toBook = Workbooks.Open(workDirectory & formatFile)

        ' �C���Ώۃt�@�C������R�s�[
        With Workbooks.Open(file)
            .Sheets("�V�[�g��").Copy Before:=toBook.Sheets(1)
            .Close
        End With
        ' �t�H�[�}�b�g�t�@�C����VBA�����s���čăt�H�[�}�b�g
        Application.Run formatFile & "!�u�a�`����"
        toBook.Close savechanges:=True
        
        ' �t�@�C���̈ړ�&���l�[��
        Name workDirectory & formatFile As toDirectory & dir(file)
    Next

    ' �m�F���b�Z�[�W�L��
    Application.DisplayAlerts = True
    ' ��ʕ`�悷��
    Application.ScreenUpdating = True
End Sub

