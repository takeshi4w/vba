Attribute VB_Name = "FSO"
Option Explicit
Private CountFiles001 As Long
Private CountFolders001 As Long
Private TotalSize001 As Double

Sub SEARCH_FOLDER001()
'-----�쐬���F2010/10/07
'-----�쐬�ҁFUchida Takeshi
'-----Update
'-----2010/10/07�F�V�K�쐬
'-----�w�肳�ꂽ�f�B���N�g���̃t�@�C���ƃT�C�Y���擾����
'-----

Dim FSO001 As New FileSystemObject
Dim RootDir001 As String
Dim OutPutSheetName001 As String

'������
CountFolders001 = 0
CountFiles001 = 0
TotalSize001 = 0

RootDir001 = Worksheets("�t�@�C��").Cells(2, 2).Value
Debug.Print RootDir001

If RootDir001 = "" Then Exit Sub

Worksheets("�t�@�C��").Select
Worksheets("�t�@�C��").Rows("10:65536").ClearContents
Set FSO001 = New FileSystemObject

' ���[�g�t�H���_����T���J�n
Call SEARCH_SUB_FOLDER001(Path001:=FSO001.GetFolder(RootDir001), Row001:=0, Column001:=0)

'OBJECT��j��
Set FSO001 = Nothing

'�I������
Worksheets("�t�@�C��").Cells(5, 2).Value = CountFolders001
Worksheets("�t�@�C��").Cells(6, 2).Value = CountFiles001
Worksheets("�t�@�C��").Cells(7, 2).Value = Format(TotalSize001, "#,##0.00") & "GB"
Worksheets("�t�@�C��").Cells(7, 2).HorizontalAlignment = xlRight
'MsgBox "�������������܂����B" & vbCr & vbCr & _
'    "�t�H���_��=" & CountFolders001 & vbCr & _
'    "�t�@�C����=" & CountFiles001 & vbCr & _
'    "�t�H���_�T�C�Y=" & Format(TotalSize001, "#,##0.00") & "GB", vbInformation

End Sub

Private Sub SEARCH_SUB_FOLDER001(ByVal Path001 As Folder, ByRef Row001 As Long, ByVal Column001 As Long)
'-----�쐬���F2010/10/07
'-----�쐬�ҁFUchida Takeshi
'-----Update
'-----2010/10/07�F�V�K�쐬
'-----�w�肳�ꂽ�f�B���N�g���̃t�@�C���ƃT�C�Y���擾����
'-----

Dim Path002 As Folder
Dim File001 As File

On Error GoTo �G���[

CountFolders001 = CountFolders001 + 1                   ' �Q�ƃt�H���_�������Z
Column001 = 1

    '�T�u�t�H���_��T��
    For Each Path002 In Path001.SubFolders
        '�ċA�Ăяo��
        Call SEARCH_SUB_FOLDER001(Path001:=Path002, Row001:=Row001, Column001:=Column001)
    Next Path002

    '�o��
    For Each File001 In Path001.Files
        Row001 = Row001 + 1
        
        '�o��
'        Worksheets("�t�@�C��").Cells(Row001 + 5, Column001).Value = "[" & Path001 & "]"
        Worksheets("�t�@�C��").Cells(Row001 + 9, Column001).Value = Path001
        Worksheets("�t�@�C��").Cells(Row001 + 9, Column001 + 1).Value = File001.Name
        Worksheets("�t�@�C��").Cells(Row001 + 9, Column001 + 2).Value = File001.DateLastModified
        Worksheets("�t�@�C��").Cells(Row001 + 9, Column001 + 3).Value = File001.Size / 1000000
        Worksheets("�t�@�C��").Cells(Row001 + 9, Column001 + 3).NumberFormatLocal = "#,##0.00"
        
        '�t�@�C�������J�E���g
        CountFiles001 = CountFiles001 + 1
        
        '�t�@�C���T�C�Y�����Z
        TotalSize001 = TotalSize001 + (File001.Size / 1000000000)
        
    Next File001
    
    'OBJECT��j��
    Set Path001 = Nothing

�G���[:
If Err.Number = 70 Then
    Resume Next
End If
Debug.Print Err.Number

End Sub
Sub SEARCH_FOLDER002()
'-----�쐬���F2010/10/07
'-----�쐬�ҁFUchida Takeshi
'-----Update
'-----2010/10/07�F�V�K�쐬
'-----�w�肳�ꂽ�f�B���N�g���̃t�@�C���ƃT�C�Y���擾����
'-----

Dim FSO001 As New FileSystemObject
Dim RootDir001 As String
Dim OutPutSheetName001 As String

'������
CountFolders001 = 0
CountFiles001 = 0
TotalSize001 = 0

RootDir001 = Worksheets("�t�H���_").Cells(2, 2).Value
Debug.Print RootDir001

If RootDir001 = "" Then Exit Sub

Worksheets("�t�H���_").Select
Worksheets("�t�H���_").Rows("9:65536").ClearContents
Set FSO001 = New FileSystemObject

' ���[�g�t�H���_����T���J�n
Call SEARCH_SUB_FOLDER002(Path001:=FSO001.GetFolder(RootDir001), Row001:=0, Column001:=0)

'OBJECT��j��
Set FSO001 = Nothing

'�I������
Worksheets("�t�H���_").Cells(5, 2).Value = CountFolders001
Worksheets("�t�H���_").Cells(6, 2).Value = Format(TotalSize001, "#,##0.00") & "GB"
Worksheets("�t�H���_").Cells(6, 2).HorizontalAlignment = xlRight
'MsgBox "�������������܂����B" & vbCr & vbCr & _
'    "�t�H���_��=" & CountFolders001 & vbCr & _
'    "�t�@�C����=" & CountFiles001 & vbCr & _
'    "�t�H���_�T�C�Y=" & Format(TotalSize001, "#,##0.00") & "GB", vbInformation

End Sub

Private Sub SEARCH_SUB_FOLDER002(ByVal Path001 As Folder, ByRef Row001 As Long, ByVal Column001 As Long)
'-----�쐬���F2010/10/07
'-----�쐬�ҁFUchida Takeshi
'-----Update
'-----2010/10/07�F�V�K�쐬
'-----�w�肳�ꂽ�f�B���N�g���̃t�@�C���ƃT�C�Y���擾����
'-----

Dim Path002 As Folder
Dim File001 As File

On Error GoTo �G���[

CountFolders001 = 0
Column001 = 1

    '�T�u�t�H���_��T��
    For Each Path002 In Path001.SubFolders
        '�o��
'        Worksheets("�t�H���_").Cells(Row001 + 5, Column001).Value = "[" & Path001 & "]"
        Row001 = Row001 + 1
        Worksheets("�t�H���_").Cells(Row001 + 8, Column001).Value = Path002
        Worksheets("�t�H���_").Cells(Row001 + 8, Column001 + 1).Value = Path002.Name
        Worksheets("�t�H���_").Cells(Row001 + 8, Column001 + 2).Value = Path002.DateLastModified
        
'        If IsError(Path002.Size) = 0 Then
            Worksheets("�t�H���_").Cells(Row001 + 8, Column001 + 3).Value = Path002.Size / 1000000
'        End If
        
        Worksheets("�t�H���_").Cells(Row001 + 8, Column001 + 3).NumberFormatLocal = "#,##0.00"
        
        '�t�H���_�����J�E���g
        CountFolders001 = CountFolders001 + 1
        
        '�t�@�C���T�C�Y�����Z
        TotalSize001 = TotalSize001 + (Path002.Size / 1000000000)
        
'        '�ċA�Ăяo��
'        Call SEARCH_SUB_FOLDER002(Path001:=Path002, Row001:=Row001, Column001:=Column001)
    Next Path002

'    '�o��
'    For Each File001 In Path001.Files
'        Row001 = Row001 + 1
'
'        '�o��
''        Worksheets("�t�H���_").Cells(Row001 + 5, Column001).Value = "[" & Path001 & "]"
'        Worksheets("�t�H���_").Cells(Row001 + 5, Column001).Value = Path001
'        Worksheets("�t�H���_").Cells(Row001 + 5, Column001 + 1).Value = File001.Name
'        Worksheets("�t�H���_").Cells(Row001 + 5, Column001 + 2).Value = File001.DateLastModified
'        Worksheets("�t�H���_").Cells(Row001 + 5, Column001 + 3).Value = File001.Size / 1000000
'        Worksheets("�t�H���_").Cells(Row001 + 5, Column001 + 3).NumberFormatLocal = "#,##0.00"
'
'        '�t�@�C�������J�E���g
'        CountFiles001 = CountFiles001 + 1
'
'        '�t�@�C���T�C�Y�����Z
'        TotalSize001 = TotalSize001 + (File001.Size / 1000000000)
'
'    Next File001
    'OBJECT��j��
    Set Path001 = Nothing

�G���[:
If Err.Number = 70 Then
    Resume Next
End If
Debug.Print Err.Number

End Sub

