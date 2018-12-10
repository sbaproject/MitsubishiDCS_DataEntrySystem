VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBankDataImport 
   Caption         =   "金融機関データ取込"
   ClientHeight    =   3345
   ClientLeft      =   2805
   ClientTop       =   1830
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5640
   Begin MSComctlLib.ProgressBar pgrProgressBar 
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "取込(&I)"
      Height          =   435
      Left            =   2100
      TabIndex        =   2
      Top             =   2580
      Width           =   1395
   End
   Begin VB.CommandButton cmdRecv 
      Caption         =   "受信(&R)"
      Height          =   435
      Left            =   540
      TabIndex        =   1
      Top             =   2580
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "終了(&X)"
      Height          =   435
      Left            =   3960
      TabIndex        =   0
      Top             =   2580
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   480
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   255
      Left            =   4020
      TabIndex        =   4
      Top             =   0
      Width           =   1395
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      Height          =   1635
      Left            =   480
      TabIndex        =   3
      Top             =   420
      Width           =   4815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "ﾌｧｲﾙ(&F)"
      Begin VB.Menu mnuEnd 
         Caption         =   "終了(&X)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "ﾍﾙﾌﾟ(&H)"
      Begin VB.Menu mnuVersion 
         Caption         =   "ﾊﾞｰｼﾞｮﾝ情報(&A)"
      End
   End
End
Attribute VB_Name = "frmBankDataImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////////////////
'//どうしても半角・全角混在のトリミングが出来ないのでこうする.
Private Type tpBank         '//金融機関
    BankCode    As String * 4   '銀行コード
    ShitenCode  As String * 3   '支店コード
    SeqCode     As String * 1   'SEQ-CODE       '銀行=':#@,=' / 支店='ｱ～ﾝ','A～Z','0～9'
    KanaName    As String * 15  '銀行名_カナ
    KanjiName   As String * 30  '銀行名_漢字
    HaitenInfo  As String * 4   '廃店情報       'Blank=営業中,1～9=廃店中
    CrLf        As String * 2   'CR + LF
End Type

Private mCaption As String
Private Const mExeMsg As String = "作業手順" & vbCrLf & vbCrLf & "　１：受信処理をします." & vbCrLf & "　２：取込処理をします." & vbCrLf & vbCrLf & "取込結果が表示されますので内容に従ってください." & vbCrLf & vbCrLf
Private mForm As New FormClass
Private mReg As New RegistryClass
Private mAbort As Boolean

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdImport_Click()
    Dim mFile As New FileClass
    
    dlgFile.DialogTitle = "ファイルを開く(" & mCaption & ")"
    dlgFile.FileName = mReg.InputFileName(mCaption)
'//LZH ファイルは解凍してからのスタートとする。
#If 1 Then
    If IsEmpty(mFile.OpenDialog(dlgFile)) Then
        Exit Sub
    End If
#Else
'// 途中までコーディングしたがやめた.....。
    If IsEmpty(mFile.OpenDialog(dlgFile, "LZHﾌｧｲﾙ (*.lzh)|*.lzh")) Then
        Exit Sub
    End If
    
    '//ファイル名をドライブ～拡張子まで分解
    Dim drv As String, path As String, file As String, ext As String
'//2006/03/13 SplitPath() にバグがあったのでコメント化：使用する時はしっかりデバックする事！
'    Call mFile.SplitPath(mReg.LzhExtractFile, drv, path, file, ext)
    '//オプション: e = Extract : 解凍
    '//パラメータ: -c 日付チェック無し
    '//            -m メッセージ抑止
    '//            -n 進捗ダイアログ非表示
    Dim ret As Integer, lzhMsg As String * 8192
    ret = Unlha(0, "e -c " & dlgFile.FileName & " " & (drv & path), lzhMsg, Len(lzhMsg))
#End If

    Dim mBank As tpBank, sql As String, SvrDate As String
    Dim updCnt As Long, insCnt As Long, delCnt As Long
    Dim fp As Integer
    Dim ms As New MouseClass
    Call ms.Start
    
    '//更新前のサーバー日付取得
    SvrDate = gdDBS.sysDate
    mReg.InputFileName(mCaption) = dlgFile.FileName
    fp = FreeFile
    Open mReg.InputFileName(mCaption) For Random Access Read As fp Len = Len(mBank)
    pgrProgressBar.Max = LOF(fp) / Len(mBank)
    '//ファイルサイズが違う場合の警告メッセージ
    If pgrProgressBar.Max <> Int(pgrProgressBar.Max) Then
        If (LOF(fp) - 1) / Len(mBank) <> Int((LOF(fp) - 1) / Len(mBank)) Then
#If 1 Then
            '/処理続行するとＤＢがおかしくなるので中止する
            Call gdDBS.MsgBox("指定されたファイル(" & mReg.InputFileName(mCaption) & ")が異常です。" & vbCrLf & vbCrLf & "処理を続行出来ません。", vbCritical + vbOKOnly, mCaption)
            Exit Sub
#Else
            If vbOK <> gdDBS.MsgBox("指定されたファイル(" & mReg.InputFileName(mCaption) & ")が異常です。" & vbCrLf & vbCrLf & "このまま続行しますか？", vbInformation + vbOKCancel + vbDefaultButton2, mCaption) Then
                Exit Sub
            End If
#End If
        End If
    End If
    
    On Error GoTo cmdImport_ClickError
    Call gdDBS.Database.BeginTrans
    'Do Until EOF(fp)   '//この構文だと最終レコードの次まで読込みしてしまう：EOF()はそういう判断
    Do While Loc(fp) < LOF(fp) / Len(mBank)
        DoEvents
        If mAbort Then
            GoTo cmdImport_ClickError
        End If
        Get fp, Loc(fp) + 1, mBank
        pgrProgressBar.Value = IIf(Loc(fp) <= pgrProgressBar.Max, Loc(fp), pgrProgressBar.Max)
        sql = "UPDATE tdBankMaster SET "
        sql = sql & " DAKJNM = '" & mFile.StrTrim(mBank.KanjiName) & "',"
        sql = sql & " DAKNNM = '" & mFile.StrTrim(mBank.KanaName) & "',"
        sql = sql & " DAHTIF = '" & mFile.StrTrim(mBank.HaitenInfo) & "',"
        sql = sql & " DAUPDT = SYSDATE"
        sql = sql & " WHERE DARKBN = '" & pGetRecordKubun(mBank.SeqCode) & "'"
        sql = sql & "   AND DABANK = '" & mFile.StrTrim(mBank.BankCode) & "'"
        sql = sql & "   AND DASITN = '" & mFile.StrTrim(mBank.ShitenCode) & "'"
        sql = sql & "   AND DASQNO = '" & mFile.StrTrim(mBank.SeqCode) & "'"
        If 0 <> gdDBS.Database.ExecuteSQL(sql) Then
            updCnt = updCnt + 1
        Else
            sql = "INSERT INTO tdBankMaster("
            sql = sql & "DARKBN,DABANK,DASITN,DASQNO,DAKNNM,DAKJNM,DAHTIF"
            sql = sql & ")VALUES("
            sql = sql & "'" & pGetRecordKubun(mBank.SeqCode) & "',"
            sql = sql & "'" & mFile.StrTrim(mBank.BankCode) & "',"
            sql = sql & "'" & mFile.StrTrim(mBank.ShitenCode) & "',"
            sql = sql & "'" & mFile.StrTrim(mBank.SeqCode) & "',"
            sql = sql & "'" & mFile.StrTrim(mBank.KanaName) & "',"
            sql = sql & "'" & mFile.StrTrim(mBank.KanjiName) & "',"
            sql = sql & "'" & mFile.StrTrim(mBank.HaitenInfo) & "'"
            sql = sql & ")"
            Call gdDBS.Database.ExecuteSQL(sql)
            insCnt = insCnt + 1
        End If
    Loop
    Close #fp
    '//更新対象でなかったレコードを削除する:必ず全件来るのが前提条件！！！
    sql = "DELETE tdBankMaster "
    sql = sql & " WHERE DAUPDT < TO_DATE('" & Format(SvrDate, "yyyy-mm-dd hh:nn:ss") & "','yyyy-mm-dd hh24:mi:ss')"
    delCnt = gdDBS.Database.ExecuteSQL(sql)
    Dim AddMsg As String
    AddMsg = "追加=" & insCnt & ":更新=" & updCnt & ":削除=" & delCnt & " 件のデータが取り込まれました。"
    lblMessage.Caption = mExeMsg & AddMsg
    Call gdDBS.AutoLogOut(mCaption, AddMsg)
    
    Call gdDBS.Database.CommitTrans
    Exit Sub
cmdImport_ClickError:
    Call gdDBS.Database.Rollback
    Call gdDBS.ErrorCheck       '//エラートラップ
'// gdDBS.ErrorCheck() の上に移動
'//    Call gdDBS.Database.Rollback
    Call gdDBS.AutoLogOut(mCaption, " エラーが発生したため取込処理は中止されました。")
End Sub

Private Function pGetRecordKubun(ByVal vCode As Variant) As Integer
    pGetRecordKubun = Abs(vCode Like "[0-9]" Or vCode Like "[A-Z]" Or vCode Like "[ｱ-ﾝ]")
End Function

Private Sub cmdRecv_Click()
    Call Shell(mReg.TransferCommand(mCaption), vbNormalFocus)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    lblMessage.Caption = mExeMsg
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mAbort = True
    Set mForm = Nothing
    Set frmBankDataImport = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub mnuEnd_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

