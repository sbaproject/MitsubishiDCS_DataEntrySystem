VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvTmpFile() As String

Private mvFileFW As Integer
Private mvFileFP As Integer
Private mvCsvData() As Variant

'// vMode = True <= フォルダのみ
Public Sub SplitPath(ByVal vFname As String, _
            Optional ByRef vDrv As Variant, _
            Optional ByRef vPath As Variant, _
            Optional ByRef vFile As Variant, _
            Optional ByRef vExt As Variant, _
            Optional ByVal vMode As Boolean = False _
        )
'//2006/03/13 バグがあるので再編
    Dim temp As String, drv As String, path As String, file As String, ext As String
    Dim ix As Integer

    '//フォルダのみの名前で呼び出されているので仮のファイル名を付加
    If True = vMode Then
        If "\" <> Right(vFname, 1) Then
            vFname = vFname & "\"
        End If
        vFname = vFname & "@@@temp.name"
    End If
    '//ファイル名の検索
    ix = InStrRev(vFname, "\")
    If ix Then
        temp = Mid(vFname, ix + 1)
        ix = InStrRev(temp, ".")
        If ix Then
            ext = Mid(temp, ix)
            file = Left(temp, ix - 1)
        Else
            file = temp
        End If
    End If
    '//ドライブ名の検索
    ix = InStr(vFname, ":")
    If ix Then
        drv = Left(vFname, ix)
    ElseIf "\\" = Left(vFname, 2) Then
        ix = InStr(Mid(vFname, 3), "\")
        drv = Left(vFname, (2 + ix) - 1)
    End If
    path = Mid(vFname, Len(drv) + 1, (Len(vFname) - Len(drv & file & ext)) - 1)
    If 0 = Len(path) Then
        path = "\"
    End If
    '/////////////////////////
    '//入力パラメータに返却
    If Not IsMissing(vDrv) And TypeName(vDrv) = "String" Then
        vDrv = drv
    End If
    If Not IsMissing(vPath) And TypeName(vPath) = "String" Then
        vPath = path
    End If
    If False = vMode Then
        If Not IsMissing(vFile) And TypeName(vFile) = "String" Then
            vFile = file
        End If
        If Not IsMissing(vExt) And TypeName(vExt) = "String" Then
            vExt = ext
        End If
    End If
End Sub

Public Function OpenCSV(fname As String) As Boolean
    On Error GoTo OpenCSVError
    mvFileFW = FreeFile
    Open fname For Input As #mvFileFW
    mvFileFP = FreeFile
    Open fname For Input As #mvFileFP
    OpenCSV = True
    'Exit Function
OpenCSVError:
    On Error GoTo 0
End Function

Public Function LineInputCSV(ByRef ColCnt As Integer) As Boolean
    On Error GoTo LineInputCSVError
    Dim tmp As String, pos As Integer
    
    Erase mvCsvData
    Line Input #mvFileFW, tmp   '１行のデータ取得
    ColCnt = 0
    If Trim(tmp) <> "" Then
        pos = -1                'カンマの存在フラグ
        Do While pos <> 0       'カンマが有れば LOOP
            ColCnt = ColCnt + 1
            ReDim Preserve mvCsvData(1 To ColCnt) As Variant
            Input #mvFileFP, mvCsvData(ColCnt)
            If Left(tmp, 1) = """" Then
                pos = InStr(Mid(tmp, 2), """,")
                tmp = Mid(tmp, pos + 3)
            Else
                pos = InStr(tmp, ",")
                tmp = Mid(tmp, pos + 1)
            End If
        Loop
        'ColCnt = ColCnt + 1
        'ReDim Preserve mvCsvData(1 To ColCnt) As Variant
        'Input #mvFileFP, mvCsvData(ColCnt)
    End If
    LineInputCSV = True
    Exit Function
LineInputCSVError:
End Function

Public Property Get GetCsvData(ary As Variant) As Variant
    On Error Resume Next
    GetCsvData = mvCsvData(ary)
End Property

Public Function MakeTempFile(Optional ByVal path As String = "C:\", Optional ByVal FileID As String = "~@") As Variant
    Dim tmpFile As String * 256
    '/////////////////////////////
    'テンポラリファイル名の生成
    If False = GetTempFileName(path, FileID, 0, tmpFile) Then
        MakeTempFile = Empty
        Exit Function
    End If
    Dim ix As Long
    Do
        ix = ix + 1
    Loop While Asc(Mid(tmpFile, ix, 1)) <> 0
    MakeTempFile = Left(tmpFile, ix - 1)
    
    ReDim Preserve mvTmpFile(0 To UBound(mvTmpFile) + 1) As String
    mvTmpFile(UBound(mvTmpFile)) = MakeTempFile
End Function

Private Sub Class_Initialize()
    ReDim mvTmpFile(0 To 0) As String
End Sub

Private Sub Class_Terminate()
    Dim ary As Integer
    On Error Resume Next
    '//mvTmpFile(0)は Dummy
    For ary = LBound(mvTmpFile) + 1 To UBound(mvTmpFile)
        Kill mvTmpFile(ary)
    Next ary
    On Error GoTo 0
    Erase mvTmpFile
End Sub

Public Property Get OpenDialog(ByVal vDlg As CommonDialog, Optional ByVal vFilter As String) As Variant
    Dim Result As Integer
    On Error GoTo OpenDialogError
OpenDialogTop:
    With vDlg
        .CancelError = True
        .Flags = cdlOFNFileMustExist + cdlOFNReadOnly + cdlOFNHideReadOnly
        If "" = vFilter Then
            .Filter = "ﾃｷｽﾄﾌｧｲﾙ (*.txt)|*.txt|すべてのﾌｧｲﾙ (*.*)|*.*"
        Else
            .Filter = vFilter & "|すべてのﾌｧｲﾙ (*.*)|*.*"
        End If
        Call .ShowOpen
        OpenDialog = Trim(.FileName)
    End With
OpenDialogError:
End Property

Public Property Get SaveDialog(ByVal vDlg As CommonDialog) As Variant
    Dim Result As Integer
    On Error GoTo SaveDialogError
SaveDialogTop:
    With vDlg
        .Flags = cdlOFNOverwritePrompt  ''//上書き確認を要求
        .CancelError = True
        .Filter = "ﾃｷｽﾄﾌｧｲﾙ(*.txt)|*.txt|すべてのﾌｧｲﾙ (*.*)|*.*"
        Call .ShowSave
        .FileName = Trim(.FileName) '//前後の空白は削除
        If .FileName <> "" Then
            If UCase(Right(.FileName, 4)) <> UCase(".txt") Then
                .FileName = .FileName & ".txt"
            End If
            SaveDialog = .FileName
        End If
    End With
SaveDialogError:
End Property

Public Function StrTrim(ByVal vData As Variant) As Variant
'//構造体のデータで末尾に Null(0) が存在すると絶対に Trim() では削除できないので作成
    Dim tmp As Variant, i As Integer
    tmp = vData
    i = Len(tmp)
    For i = Len(tmp) To 1 Step -1
        If Asc(Mid(tmp, i, 1)) <> 0 Then
            Exit For
        End If
        Mid(tmp, i, 1) = " "
    Next i
    StrTrim = Trim(tmp)
End Function


