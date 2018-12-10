Attribute VB_Name = "MainModule"
Option Explicit
Public gdDBS  As DatabaseClass
Public gdForm As Form
Public gdFormSub As Form         '//子供の画面が存在する？

'//2006/03/10 保護者マスタの取込ユーザーＩＤ
Public Const gcImportUserName As String = "PUNCH_IMPORT"
Public Const gcTsuchoKigoMinLen As Integer = 3
Public Const gcTsuchoBangoMinLen As Integer = 7
Public Const gcFurikaeImportToDelete As String = "D"    '//予定廃棄
Public Const gcImportHogoshaUser    As String = "$KZ_IMP"   '//口座振替依頼書・取込ユーザー
Public Const gcFurikaeImportToMaster As String = "M"    '//マスター反映

Public Enum eBankKubun
    KinnyuuKikan = 0
    YuubinKyoku = 1
End Enum

Public Enum eBankRecordKubun
    Bank = 0
    Shiten = 1
End Enum

Public Enum eBankYokinShubetsu
    Dummy = 0
    Futsuu = 1
    Touza = 2
End Enum

Public Enum eShoriKubun
    Add = 0
    Edit = 1
    Delete = 2
    Refer = 3   '//2012/12/07 参照のオプションボタン追加
End Enum

Public Enum eKouFuriKubun
    YoteiDB = 0         '//予定ＤＢ作成
    YoteiText = 1       '//予定テキスト作成
    YoteiImport = 2     '//予定データ取込
    SeikyuText = 3      '//請求テキスト作成
End Enum

#If 0 Then
'//全銀マスタ取り込み用
Declare Function Unlha Lib "Unlha32.DLL" (ByVal hWnd As Integer, ByVal szCmdline As String, ByVal szOutPutMsg As String, ByVal dwSize As Long) As Integer
#End If

'//2014/06/11 リストから選べる内容をチェックするために定数化
Public Const cKAIYAKU_DATA As String = "保護者マスタは解約状態です."
Public Const cEXISTS_DATA As String = "保護者マスタに既に存在します."

Sub Main()
    Dim mFile As New FileClass, path As String, drv As String
    Set gdDBS = New DatabaseClass
    Call frmMainMenu.Show
End Sub

Sub gkAllEnd()
    Set gdDBS = Nothing
    Set gdForm = Nothing
    End
End Sub


'////////////////////////////////////////////////////////////////////
'//EXE のプログラム配下に \Backup フォルダを作成してバックアップする
Public Function gBackupTextData(vFileName As String) As Boolean
    Dim mFile As New FileClass
    Dim dstPath As String, dstDrv As String
    Dim dstFile As String, dstExt As String
    Call mFile.SplitPath(App.path, vDrv:=dstDrv, vPath:=dstPath, vMode:=True)
    Call mFile.SplitPath(vFileName, vFile:=dstFile, vExt:=dstExt)
    On Error Resume Next
    If "" = Dir(dstDrv & dstPath & "\Backup\") Then
        Call MkDir(dstDrv & dstPath & "\Backup")
    End If
    Call FileCopy(vFileName, dstDrv & dstPath & "\Backup\" & Format(Now(), "yyyymmdd.hhnnss") & "." & dstFile & dstExt)
End Function

