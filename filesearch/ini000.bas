Attribute VB_Name = "ini000"
Public Sub LogWrite001(ProgName001, StartEnd001, Notes001, Procedure001, LogSQL001)
'-----�쐬���F2008/10/30 Uchida Takeshi
'-----Update
'-----2008/10/30�F�V�K�쐬
'-----2008/11/25�F_tbl_��{�ݒ�Ƃ̘A�g��ǉ�
'-----���O�̏�������
'-----

'Dim FSO As New FileSystemObject
Dim FSO As New FileSystemObject
Dim TextFile As TextStream
Dim buf As String
Dim logfileNAME001 As String
Dim Dir001 As Variant '�f�B���N�g�����i�[����

On Error GoTo �G���[

Dir001 = DLookup("[���e]", "[_tbl_��{�ݒ�]", "[�p�r] = '���O�t�@�C���̏o�͐�'")
If Dir001 = "CurrentProject.Path" Then
    Dir001 = CurrentProject.Path
End If

logfileNAME001 = DLookup("[���e]", "[_tbl_��{�ݒ�]", "[�p�r] = '���O�t�@�C����'")

'Log�������݁i�J�n�j
Set FSO = CreateObject("Scripting.FileSystemObject")
Set TextFile = FSO.OpenTextFile(Dir001 & "\" & logfileNAME001, ForAppending, True)

With TextFile
    .Write (Now & ",")                      '����
    .Write (CurrentProject.Path & "\")
    .Write (CurrentProject.Name & ",")      '�t�@�C���̃f�B���N�g��
    .Write (StartEnd001 & ",")              '�J�n/�I��
    .Write (Notes001 & ",")                 '��Ɠ��e
    .Write (Procedure001 & ",")             '�v���V�[�W����
    .Write (Chr(34) & LogSQL001 & Chr(34) & ",")                '���s���ꂽSQL��
    .Write (ProgName001)                    '�v���O������
    .WriteLine
    .Close
End With

�G���[:
If Err.Number = 1004 Then
    Resume
'Else
'    MsgBox Err.Number & " : " & Err.Description
End If


End Sub

