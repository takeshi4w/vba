Attribute VB_Name = "FSO"
Option Explicit
Private CountFiles001 As Long
Private CountFolders001 As Long
Private TotalSize001 As Double

Sub SEARCH_FOLDER001()
'-----作成日：2010/10/07
'-----作成者：Uchida Takeshi
'-----Update
'-----2010/10/07：新規作成
'-----指定されたディレクトリのファイルとサイズを取得する
'-----

Dim FSO001 As New FileSystemObject
Dim RootDir001 As String
Dim OutPutSheetName001 As String

'初期化
CountFolders001 = 0
CountFiles001 = 0
TotalSize001 = 0

RootDir001 = Worksheets("ファイル").Cells(2, 2).Value
Debug.Print RootDir001

If RootDir001 = "" Then Exit Sub

Worksheets("ファイル").Select
Worksheets("ファイル").Rows("10:65536").ClearContents
Set FSO001 = New FileSystemObject

' ルートフォルダから探索開始
Call SEARCH_SUB_FOLDER001(Path001:=FSO001.GetFolder(RootDir001), Row001:=0, Column001:=0)

'OBJECTを破棄
Set FSO001 = Nothing

'終了処理
Worksheets("ファイル").Cells(5, 2).Value = CountFolders001
Worksheets("ファイル").Cells(6, 2).Value = CountFiles001
Worksheets("ファイル").Cells(7, 2).Value = Format(TotalSize001, "#,##0.00") & "GB"
Worksheets("ファイル").Cells(7, 2).HorizontalAlignment = xlRight
'MsgBox "処理が完了しました。" & vbCr & vbCr & _
'    "フォルダ数=" & CountFolders001 & vbCr & _
'    "ファイル数=" & CountFiles001 & vbCr & _
'    "フォルダサイズ=" & Format(TotalSize001, "#,##0.00") & "GB", vbInformation

End Sub

Private Sub SEARCH_SUB_FOLDER001(ByVal Path001 As Folder, ByRef Row001 As Long, ByVal Column001 As Long)
'-----作成日：2010/10/07
'-----作成者：Uchida Takeshi
'-----Update
'-----2010/10/07：新規作成
'-----指定されたディレクトリのファイルとサイズを取得する
'-----

Dim Path002 As Folder
Dim File001 As File

On Error GoTo エラー

CountFolders001 = CountFolders001 + 1                   ' 参照フォルダ数を加算
Column001 = 1

    'サブフォルダを探索
    For Each Path002 In Path001.SubFolders
        '再帰呼び出し
        Call SEARCH_SUB_FOLDER001(Path001:=Path002, Row001:=Row001, Column001:=Column001)
    Next Path002

    '出力
    For Each File001 In Path001.Files
        Row001 = Row001 + 1
        
        '出力
'        Worksheets("ファイル").Cells(Row001 + 5, Column001).Value = "[" & Path001 & "]"
        Worksheets("ファイル").Cells(Row001 + 9, Column001).Value = Path001
        Worksheets("ファイル").Cells(Row001 + 9, Column001 + 1).Value = File001.Name
        Worksheets("ファイル").Cells(Row001 + 9, Column001 + 2).Value = File001.DateLastModified
        Worksheets("ファイル").Cells(Row001 + 9, Column001 + 3).Value = File001.Size / 1000000
        Worksheets("ファイル").Cells(Row001 + 9, Column001 + 3).NumberFormatLocal = "#,##0.00"
        
        'ファイル数をカウント
        CountFiles001 = CountFiles001 + 1
        
        'ファイルサイズを加算
        TotalSize001 = TotalSize001 + (File001.Size / 1000000000)
        
    Next File001
    
    'OBJECTを破棄
    Set Path001 = Nothing

エラー:
If Err.Number = 70 Then
    Resume Next
End If
Debug.Print Err.Number

End Sub
Sub SEARCH_FOLDER002()
'-----作成日：2010/10/07
'-----作成者：Uchida Takeshi
'-----Update
'-----2010/10/07：新規作成
'-----指定されたディレクトリのファイルとサイズを取得する
'-----

Dim FSO001 As New FileSystemObject
Dim RootDir001 As String
Dim OutPutSheetName001 As String

'初期化
CountFolders001 = 0
CountFiles001 = 0
TotalSize001 = 0

RootDir001 = Worksheets("フォルダ").Cells(2, 2).Value
Debug.Print RootDir001

If RootDir001 = "" Then Exit Sub

Worksheets("フォルダ").Select
Worksheets("フォルダ").Rows("9:65536").ClearContents
Set FSO001 = New FileSystemObject

' ルートフォルダから探索開始
Call SEARCH_SUB_FOLDER002(Path001:=FSO001.GetFolder(RootDir001), Row001:=0, Column001:=0)

'OBJECTを破棄
Set FSO001 = Nothing

'終了処理
Worksheets("フォルダ").Cells(5, 2).Value = CountFolders001
Worksheets("フォルダ").Cells(6, 2).Value = Format(TotalSize001, "#,##0.00") & "GB"
Worksheets("フォルダ").Cells(6, 2).HorizontalAlignment = xlRight
'MsgBox "処理が完了しました。" & vbCr & vbCr & _
'    "フォルダ数=" & CountFolders001 & vbCr & _
'    "ファイル数=" & CountFiles001 & vbCr & _
'    "フォルダサイズ=" & Format(TotalSize001, "#,##0.00") & "GB", vbInformation

End Sub

Private Sub SEARCH_SUB_FOLDER002(ByVal Path001 As Folder, ByRef Row001 As Long, ByVal Column001 As Long)
'-----作成日：2010/10/07
'-----作成者：Uchida Takeshi
'-----Update
'-----2010/10/07：新規作成
'-----指定されたディレクトリのファイルとサイズを取得する
'-----

Dim Path002 As Folder
Dim File001 As File

On Error GoTo エラー

CountFolders001 = 0
Column001 = 1

    'サブフォルダを探索
    For Each Path002 In Path001.SubFolders
        '出力
'        Worksheets("フォルダ").Cells(Row001 + 5, Column001).Value = "[" & Path001 & "]"
        Row001 = Row001 + 1
        Worksheets("フォルダ").Cells(Row001 + 8, Column001).Value = Path002
        Worksheets("フォルダ").Cells(Row001 + 8, Column001 + 1).Value = Path002.Name
        Worksheets("フォルダ").Cells(Row001 + 8, Column001 + 2).Value = Path002.DateLastModified
        
'        If IsError(Path002.Size) = 0 Then
            Worksheets("フォルダ").Cells(Row001 + 8, Column001 + 3).Value = Path002.Size / 1000000
'        End If
        
        Worksheets("フォルダ").Cells(Row001 + 8, Column001 + 3).NumberFormatLocal = "#,##0.00"
        
        'フォルダ数をカウント
        CountFolders001 = CountFolders001 + 1
        
        'ファイルサイズを加算
        TotalSize001 = TotalSize001 + (Path002.Size / 1000000000)
        
'        '再帰呼び出し
'        Call SEARCH_SUB_FOLDER002(Path001:=Path002, Row001:=Row001, Column001:=Column001)
    Next Path002

'    '出力
'    For Each File001 In Path001.Files
'        Row001 = Row001 + 1
'
'        '出力
''        Worksheets("フォルダ").Cells(Row001 + 5, Column001).Value = "[" & Path001 & "]"
'        Worksheets("フォルダ").Cells(Row001 + 5, Column001).Value = Path001
'        Worksheets("フォルダ").Cells(Row001 + 5, Column001 + 1).Value = File001.Name
'        Worksheets("フォルダ").Cells(Row001 + 5, Column001 + 2).Value = File001.DateLastModified
'        Worksheets("フォルダ").Cells(Row001 + 5, Column001 + 3).Value = File001.Size / 1000000
'        Worksheets("フォルダ").Cells(Row001 + 5, Column001 + 3).NumberFormatLocal = "#,##0.00"
'
'        'ファイル数をカウント
'        CountFiles001 = CountFiles001 + 1
'
'        'ファイルサイズを加算
'        TotalSize001 = TotalSize001 + (File001.Size / 1000000000)
'
'    Next File001
    'OBJECTを破棄
    Set Path001 = Nothing

エラー:
If Err.Number = 70 Then
    Resume Next
End If
Debug.Print Err.Number

End Sub

