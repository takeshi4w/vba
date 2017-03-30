Attribute VB_Name = "ini000"
Public Sub LogWrite001(ProgName001, StartEnd001, Notes001, Procedure001, LogSQL001)
'-----作成日：2008/10/30 Uchida Takeshi
'-----Update
'-----2008/10/30：新規作成
'-----2008/11/25：_tbl_基本設定との連携を追加
'-----ログの書き込み
'-----

'Dim FSO As New FileSystemObject
Dim FSO As New FileSystemObject
Dim TextFile As TextStream
Dim buf As String
Dim logfileNAME001 As String
Dim Dir001 As Variant 'ディレクトリを格納する

On Error GoTo エラー

Dir001 = DLookup("[内容]", "[_tbl_基本設定]", "[用途] = 'ログファイルの出力先'")
If Dir001 = "CurrentProject.Path" Then
    Dir001 = CurrentProject.Path
End If

logfileNAME001 = DLookup("[内容]", "[_tbl_基本設定]", "[用途] = 'ログファイル名'")

'Log書き込み（開始）
Set FSO = CreateObject("Scripting.FileSystemObject")
Set TextFile = FSO.OpenTextFile(Dir001 & "\" & logfileNAME001, ForAppending, True)

With TextFile
    .Write (Now & ",")                      '時刻
    .Write (CurrentProject.Path & "\")
    .Write (CurrentProject.Name & ",")      'ファイルのディレクトリ
    .Write (StartEnd001 & ",")              '開始/終了
    .Write (Notes001 & ",")                 '作業内容
    .Write (Procedure001 & ",")             'プロシージャ名
    .Write (Chr(34) & LogSQL001 & Chr(34) & ",")                '発行されたSQL文
    .Write (ProgName001)                    'プログラム名
    .WriteLine
    .Close
End With

エラー:
If Err.Number = 1004 Then
    Resume
'Else
'    MsgBox Err.Number & " : " & Err.Description
End If


End Sub

