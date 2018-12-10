VERSION 5.00
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKouzaFurikaeExport 
   Caption         =   "口座振替データ作成"
   ClientHeight    =   4035
   ClientLeft      =   2295
   ClientTop       =   2235
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6735
   Begin MSComctlLib.ProgressBar pgbRecord 
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   2700
      Visible         =   0   'False
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdMakeText 
      Caption         =   "テキスト作成(&T)"
      Height          =   435
      Left            =   1620
      TabIndex        =   3
      Top             =   3180
      Width           =   1395
   End
   Begin VB.ComboBox cboFurikaeBi 
      Height          =   300
      Left            =   3060
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   310
      Width           =   1275
   End
   Begin VB.CheckBox chkJisseki 
      BackColor       =   &H000000FF&
      Caption         =   "1 = 確定(WAOは確定のみ)"
      Height          =   435
      Left            =   180
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   60
      Value           =   2  '無効
      Width           =   1815
   End
   Begin VB.CommandButton cmdMakeDB 
      Caption         =   "ＤＢ作成(&D)"
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   3180
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "終了(&X)"
      Default         =   -1  'True
      Height          =   435
      Left            =   5280
      TabIndex        =   4
      Top             =   3180
      Width           =   1335
   End
   Begin imDate6Ctl.imDate txtFurikaeBi 
      Height          =   285
      Left            =   180
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   540
      Visible         =   0   'False
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   503
      Calendar        =   "口座振替データ作成.frx":0000
      Caption         =   "口座振替データ作成.frx":0186
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "口座振替データ作成.frx":01F4
      Keys            =   "口座振替データ作成.frx":0212
      MouseIcon       =   "口座振替データ作成.frx":0270
      Spin            =   "口座振替データ作成.frx":028C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   255
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "yyyy/mm/dd"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy/mm/dd"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   73050
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   " "
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "    /  /  "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   -2
      CenturyMode     =   0
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   4740
      Top             =   3180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label1"
      Height          =   315
      Left            =   5220
      TabIndex        =   8
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label8 
      Caption         =   "口座振替日"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   360
      Width           =   915
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      Height          =   2175
      Left            =   420
      TabIndex        =   1
      Top             =   960
      Width           =   6015
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
Attribute VB_Name = "frmKouzaFurikaeExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCaption As String
Private Const mExeMsg As String = "作成処理をします." & vbCrLf & vbCrLf & "作成結果が表示されますので内容に従ってください." & vbCrLf & vbCrLf
Private mForm As New FormClass
Private mAbort As Boolean

Private Enum eCheckButton
    Yotei = 0
    Kakutei = 1
    Mukou = 2
End Enum

Private Sub cboFurikaeBi_Click()
    txtFurikaeBi.Text = cboFurikaeBi.Text
    Dim sql As String, dyn As Object
    sql = "SELECT FASQNO"
    sql = sql & " FROM tfFurikaeYoteiData"
    sql = sql & " WHERE FASQNO = " & txtFurikaeBi.Number
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    cmdMakeText.Enabled = (dyn.EOF = False) 'データがあればテキスト作成可能
    Call dyn.Close
End Sub

Private Sub chkJisseki_Click()
    '//実績の時は日付は変更不可：最終のデータで作成する
    'txtFurikaeBi.Enabled = chkJisseki.Value = eCheckButton.Yotei
    'cboFurikaeBi.Enabled = chkJisseki.Value = eCheckButton.Yotei
'//2004/04/13 請求時にＤＢ作成を有効にする＆テキスト作成・送信を無効にする：ＤＢ作成後有効に！
'//    cmdMakeDB.Enabled = chkJisseki.Value = eCheckButton.Yotei
    cmdMakeText.Enabled = (chkJisseki.Value = eCheckButton.Yotei)
    'cmdSend.Enabled = chkJisseki.Value = eCheckButton.Yotei

'    Dim sql As String, dyn As OraDynaset, MaxDay As Variant
    Dim sql As String, dyn As Object, MaxDay As Variant
    sql = "SELECT FASQNO,TO_CHAR(TO_DATE(FASQNO,'YYYYMMDD'),'YYYY/MM/DD') AS FaDate"
    sql = sql & " FROM tfFurikaeYoteiData"
    sql = sql & " GROUP BY FASQNO"
    sql = sql & " ORDER BY FASQNO"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    Call cboFurikaeBi.Clear
    Do Until dyn.EOF()
        Call cboFurikaeBi.AddItem(dyn.Fields("FaDate"))
        cboFurikaeBi.ItemData(cboFurikaeBi.NewIndex) = dyn.Fields("FASQNO")
        MaxDay = dyn.Fields("FASQNO")
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    '//予定の時は基本情報の次回振替日を追加
'    If chkJisseki.Value = eCheckButton.Yotei Then
        sql = "SELECT AANXKZ,TO_CHAR(TO_DATE(AANXKZ,'YYYYMMDD'),'YYYY/MM/DD') AS AaDate"
        sql = sql & " FROM taSystemInformation"
        sql = sql & " WHERE AASKEY = 'SYSTEM'"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
        If Not dyn.EOF() Then
            '//振替予定データの最終口座振替日より大きい時のみ
            If MaxDay < dyn.Fields("AANXKZ") Then
                Call cboFurikaeBi.AddItem(dyn.Fields("AaDate"))
                cboFurikaeBi.ItemData(cboFurikaeBi.NewIndex) = dyn.Fields("AANXKZ")
            End If
        End If
'    End If
    If cboFurikaeBi.ListCount Then
        cboFurikaeBi.ListIndex = cboFurikaeBi.ListCount - 1
    End If
    Dim ary As Variant
    ary = Array("(予定)", "(請求)")
    mCaption = Left(mCaption, IIf(InStr(mCaption, "("), InStr(mCaption, "(") - 1, Len(mCaption)))
    Me.Caption = Left(Me.Caption, IIf(InStr(Me.Caption, mCaption), InStr(Me.Caption, mCaption) - 1, Len(Me.Caption)))
    mCaption = mCaption & ary(chkJisseki.Value)
    Me.Caption = Me.Caption & mCaption
'//2004/04/13 請求時にＤＢ作成を有効にする＆テキスト作成・送信を無効にする：ＤＢ作成後有効に！
'//    cmdMakeText.Enabled = cboFurikaeBi.ListCount > 0
End Sub

Private Sub cmdEnd_Click()
    Unload Me
End Sub

#Const cSPEEDUP = True

Private Sub cmdMakeDB_Click()
    On Error GoTo cmdExport_ClickError
'    Dim sql As String, dyn As OraDynaset
    Dim sql As String, dyn As Object
    Dim reg As New RegistryClass
    
'//2003/01/30 過去データを再作成できなくする
    If txtFurikaeBi.Text < gdDBS.sysDate("YYYY/MM/DD") Then
        Call MsgBox("ＤＢ作成をしようとしている日付は過去の日付です." & vbCrLf & vbCrLf & _
                    "過去日付データは作成できません." & vbCrLf & vbCrLf & _
                    "サーバー(" & reg.DbDatabaseName & ")日付 = " & gdDBS.sysDate("YYYY/MM/DD"), vbInformation + vbOKOnly, mCaption)
        Exit Sub
    End If
'//2004/04/13 複数月の予定データは作成できないように制御する。
'// If cboFurikaeBi.ListCount > 1 Then
    If cboFurikaeBi.ListIndex > 0 Then
        Call MsgBox("複数月のＤＢ作成(予定)は出来ません." & vbCrLf & vbCrLf & _
                    "先に振替予定表の累積処理を実行してください." _
                    , vbInformation + vbOKOnly, mCaption)
        Exit Sub
    End If
'// End If
    
    '//同一契約者が複数件あると保護者がその件数分の結果が返るので ==> DISTINCT
    sql = "SELECT DISTINCT a.ABITCD,c.* "
    sql = sql & " FROM taItakushaMaster     a,"
    sql = sql & "      tbKeiyakushaMaster   b,"
    '//基本は保護者マスター
    sql = sql & "      tcHogoshaMaster      c "
    sql = sql & " WHERE ABITKB = BAITKB"
    sql = sql & "   AND BAITKB = CAITKB"
    sql = sql & "   AND BAKYCD = CAKYCD"
'//2002/12/10 教室区分(??KSCD)は使用しない
'//    sql = sql & "   AND BAKSCD = CAKSCD"
    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN CAFKST AND CAFKED"
'//2003/02/03 解約フラグ参照追加
    sql = sql & "   AND NVL(BAKYFG,0) = 0"  '//契約者は解約していない
    sql = sql & "   AND NVL(CAKYFG,0) = 0"  '//保護者は解約していない
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If dyn.EOF Then
        Call MsgBox(txtFurikaeBi.Text & " に該当するデータはありません.", vbInformation + vbOKOnly, mCaption)
        Exit Sub
    End If
'//2003/02/03 システムのフラグを参照してしようと思ったが複数日のデータが有ると出来ないのでやめた.
'//    If gdDBS.SystemUpdate("AAUPD2").Value <> 0 Then
'//        Call MsgBox(txtFurikaeBi.Text & " に該当するデータはありません.", vbInformation + vbOKOnly, mCaption)
'//        Exit Sub
'//    End If
    
    Dim ms As New MouseClass
    Call ms.Start
    
'//2003/01/31 新規エントリーデータ判断用システム記憶日
    Dim NewEntryStartDate As String, ReMake As Boolean
    NewEntryStartDate = Format(gdDBS.SystemUpdate("AANWDT"), "yyyy/mm/dd hh:nn:ss")
    
    Call gdDBS.Database.BeginTrans
    
    '//関連テーブルロック：2004/04/13 本当にロックできるの？
    Call gdDBS.Database.ExecuteSQL("Lock Table tbKeiyakushaMaster IN EXCLUSIVE MODE NOWAIT")
    Call gdDBS.Database.ExecuteSQL("Lock Table tcHogoshaMaster    IN EXCLUSIVE MODE NOWAIT")
    Call gdDBS.Database.ExecuteSQL("Lock Table tfFurikaeYoteiData IN EXCLUSIVE MODE NOWAIT")
    Call gdDBS.Database.ExecuteSQL("Lock Table tfFurikaeYoteiTran IN EXCLUSIVE MODE NOWAIT")
    
    sql = "DELETE tfFurikaeYoteiData "
    sql = sql & " WHERE FASQNO = '" & txtFurikaeBi.Number & "'"
    If 0 <> gdDBS.Database.ExecuteSQL(sql) Then
        If vbYes <> MsgBox(txtFurikaeBi.Text & " のデータは既に存在します." & vbCrLf & vbCrLf & "再度作成しなおますか？", vbInformation + vbDefaultButton3 + vbYesNoCancel, Me.Caption) Then
            GoTo cmdExport_ClickError
        End If
'//2003/02/03 再作成時は予定作成日を更新しない
        ReMake = True
    End If
    Dim cnt As Long

Debug.Print "start= " & Now

'////////////////////////////////////////////
'//2012/07/11 スピードアップ改善：ここから
#If cSPEEDUP = False Then
'''    Do Until dyn.EOF
'''        DoEvents
'''        If mAbort Then
'''            GoTo cmdExport_ClickError
'''        End If
'''        cnt = cnt + 1
'''        '//振替予定データに追加
'''        sql = "INSERT INTO tfFurikaeYoteiData VALUES("
''''//2003/01/31 Dynaset を Object で定義すると .Value 句を付加しないと Error=5 になる.
'''        sql = sql & "'" & dyn.Fields("CAITKB").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAKYCD").Value & "',"
'''        'sql = sql & "'" & dyn.Fields("CAKSCD").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAHGCD").Value & "',"
'''        sql = sql & "'" & txtFurikaeBi.Number & "',"
'''        sql = sql & "'" & dyn.Fields("CAKKBN").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CABANK").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CASITN").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAKZSB").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAKZNO").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAYBTK").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAYBTN").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAKZNM").Value & "',"
'''        sql = sql & "0,0,0,"                                  '//請求金額・変更後金額・解約フラグ
'''        sql = sql & Abs(IsNull(dyn.Fields("CANWDT").Value)) & ","
'''        sql = sql & "'" & dyn.Fields("CAFKST").Value & "',"
'''        sql = sql & eKouFuriKubun.YoteiDB & ","
'''        sql = sql & "'" & gdDBS.LoginUserName & "',"
'''        sql = sql & "SYSDATE"
'''        sql = sql & ")"
'''        Call gdDBS.Database.ExecuteSQL(sql)
'''        Call dyn.MoveNext
'''    Loop
#Else   'cSPEEDUP = False Then
    
    cnt = dyn.RecordCount
    
    sql = "INSERT INTO tfFurikaeYoteiData "
    sql = sql & "SELECT DISTINCT "
    sql = sql & "CAITKB,"
    sql = sql & "CAKYCD,"
    'sql = sql & "CAKSCD,"
    sql = sql & "CAHGCD,"
    sql = sql & txtFurikaeBi.Number & ","
    sql = sql & "CAKKBN,"
    sql = sql & "CABANK,"
    sql = sql & "CASITN,"
    sql = sql & "CAKZSB,"
    sql = sql & "CAKZNO,"
    sql = sql & "CAYBTK,"
    sql = sql & "CAYBTN,"
    sql = sql & "CAKZNM,"
    sql = sql & "0,0,0,"
    sql = sql & "(case when CANWDT is null then 1 else 0 end) CANWDT,"
    sql = sql & "CAFKST,"
    sql = sql & eKouFuriKubun.YoteiDB & ","
    sql = sql & "'" & gdDBS.LoginUserName & "',"
    sql = sql & "SYSDATE"
    sql = sql & " FROM taItakushaMaster     a,"
    sql = sql & "      tbKeiyakushaMaster   b,"
    '//基本は保護者マスター
    sql = sql & "      tcHogoshaMaster      c "
    sql = sql & " WHERE ABITKB = BAITKB"
    sql = sql & "   AND BAITKB = CAITKB"
    sql = sql & "   AND BAKYCD = CAKYCD"
    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN CAFKST AND CAFKED"
    sql = sql & "   AND NVL(BAKYFG,0) = 0"  '//契約者は解約していない
    sql = sql & "   AND NVL(CAKYFG,0) = 0"  '//保護者は解約していない
    Dim insCnt As Long
    insCnt = gdDBS.Database.ExecuteSQL(sql)
    If insCnt <> cnt Then
        Call err.Raise(-1, "cmdMakeDB", "ＤＢ作成は失敗しました.")
    End If
#End If     'cSPEEDUP = False Then
'//2012/07/11 スピードアップ改善：ここまで
'////////////////////////////////////////////

Debug.Print "  end= " & Now
    
    Call dyn.Close
    '//実行更新フラグ設定：この関数は予定のみ実行可能
'//2003/02/03 再作成時は予定作成日を更新しない
    If ReMake = False Then
        gdDBS.SystemUpdate("AAUPD1") = 1
    End If
    
'//2004/05/17 詳細を関数化
    Call pNormalEndMessage(ReMake, cnt, NewEntryStartDate)
    
    Call gdDBS.Database.CommitTrans
    
'//2004/04/13 請求時にＤＢ作成を有効にする＆テキスト作成・送信を無効にする：ＤＢ作成後有効に！
    cmdMakeText.Enabled = True
'    cmdSend.Enabled = True
    Exit Sub
cmdExport_ClickError:
    Call gdDBS.Database.Rollback
    Call gdDBS.ErrorCheck(gdDBS.Database)       '//エラートラップ
'// gdDBS.ErrorCheck() の上に移動
'//    Call gdDBS.Database.Rollback
End Sub

Private Sub pNormalEndMessage(ByVal vRemake As Boolean, vCnt As Long, ByVal vNewEntryStartDate As Variant)
'//2004/05/17 詳細を関数化
    Dim sql As String, dyn As Object

'//2004/04/26 新規件数&０円件数＆合計金額の追加
'//2004/05/17 総０円カウントを新規の０円に変更
    Dim NewCnt As Long, NewZero As Long, TotalGaku As Currency ', ZeroCnt As Long
'//2006/04/25 新規解約件数カウント追加
    Dim CanCnt As Long, JsNewCnt As Long
    sql = "SELECT " & vbCrLf
    sql = sql & " SUM(NewCnt)   NewCnt," & vbCrLf                   '//新規件数
    sql = sql & " SUM(NewZero)  NewZero," & vbCrLf                  '//新規０円件数
    'sql = sql & " SUM(ZeroCnt)  ZeroCnt," & vbCrLf                  '//０円総件数
    sql = sql & " SUM(TtlGaku)  TtlGaku," & vbCrLf                  '//総請求金額
    sql = sql & " SUM(CanCnt)   CanCnt ," & vbCrLf                  '//2006/04/25 新規解約
    sql = sql & " SUM(JsNewCnt) JsNewCnt" & vbCrLf                  '//2006/08/10 実際の新規件数
    sql = sql & " FROM (" & vbCrLf
    sql = sql & " SELECT " & vbCrLf
    sql = sql & " SUM(NVL(FANWCD,0)) AS NewCnt," & vbCrLf                        '//新規件数
    sql = sql & " SUM(DECODE(FASKGK,0,NVL(FANWCD,0),0)) AS NewZero," & vbCrLf    '//新規０円件数
    'sql = sql & " SUM(DECODE(FASKGK,0,1,0)) AS ZeroCnt," & vbCrLf                '//０円総件数
    sql = sql & " SUM(NVL(FASKGK,0)) AS TtlGaku," & vbCrLf                       '//総請求金額
    sql = sql & " 0 AS CanCnt,  " & vbCrLf                                       '//2006/04/25 新規解約
    sql = sql & " 0 AS JsNewCnt " & vbCrLf                                       '//2006/08/10 実際の新規件数
    sql = sql & " FROM tfFurikaeYoteiData " & vbCrLf
    sql = sql & " WHERE FASQNO = '" & txtFurikaeBi.Number & "'" & vbCrLf
    sql = sql & " UNION " & vbCrLf
    sql = sql & " SELECT " & vbCrLf
    sql = sql & " 0 AS NewCnt," & vbCrLf                 '//新規件数
    sql = sql & " 0 AS NewZero," & vbCrLf                '//新規０円件数
    'sql = sql & " 0 AS ZeroCnt," & vbCrLf                '//０円総件数
    sql = sql & " 0 AS TtlGaku," & vbCrLf                '//総請求金額
    sql = sql & " SUM(DECODE(NVL(CAKYFG,0),0,0,1)) AS CanCnt," & vbCrLf          '//2006/04/25 新規解約
    sql = sql & " COUNT(*) AS JsNewCnt " & vbCrLf        '//2006/08/10 実際の新規件数
    sql = sql & " FROM tcHogoshaMaster " & vbCrLf
    sql = sql & " WHERE CAADDT >= TO_DATE('" & vNewEntryStartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
'//2006/08/10 実際の新規件数の為にコメント化
'//    sql = sql & "   AND NVL(CAKYFG,0) <> 0" & vbCrLf
    sql = sql & ")"
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
    If Not dyn.EOF Then
        NewCnt = dyn.Fields("NewCnt").Value
        NewZero = dyn.Fields("NewZero").Value
        'ZeroCnt = dyn.Fields("ZeroCnt").Value
        TotalGaku = dyn.Fields("TtlGaku").Value
'//2006/04/25 新規解約件数カウント追加
        CanCnt = dyn.Fields("CanCnt").Value
'//2006/08/10 実際の新規件数の為にコメント化
        JsNewCnt = dyn.Fields("JsNewCnt").Value
    End If
    Call dyn.Close
    
'//2004/04/26 新規件数&０円件数＆合計金額の追加
'//2004/05/17 総０円カウントを新規の０円に変更
'//2006/04/25 新規解約件数カウント追加
    Call gdDBS.AutoLogOut(mCaption, "ＤＢ" & IIf(vRemake = True, "再", "新規") & "作成(" & _
                "口座振替日=[" & txtFurikaeBi.Text & "] : 新規データ対象登録日=[" & Format(vNewEntryStartDate, "yyyy/mm/dd hh:nn:ss") & "] : 作成件数=" & vCnt & " 件)" & _
                " <新規件数=" & NewCnt & ">")
'//2004/04/26 新規件数&０円件数＆合計金額の追加
'//2004/05/17 総０円カウントを新規の０円に変更
'//2006/04/25 新規解約件数カウント追加
    lblMessage.Caption = vCnt & " 件のデータが作成されました。" & vbCrLf & vbCrLf & _
                        "<< 新規件数の詳細 >>" & vbCrLf & _
                        "　請求件数 = " & NewCnt - NewZero & vbCrLf & _
                        "-------------------" & vbCrLf & _
                        "　新規件数 = " & NewCnt & vbCrLf & _
                        "===================" & vbCrLf & _
                        "<<  総件数 = " & vCnt & " >>"
                        
    Call MsgBox(IIf(vRemake = True, "再", "新規") & "作成は正常終了しました。" & vbCrLf & vbCrLf & "出力メッセージの内容を確認して下さい。", vbInformation, mCaption)
End Sub

Private Sub cmdMakeText_Click()
    On Error GoTo cmdExport_ClickError
'    Dim sql As String, dyn As OraDynaset
    Dim sql As String, dyn As Object
    
    sql = "SELECT a.ABITCD,c.*,SUBSTR(LPAD(FAFKST,8,'0'),1,6) FAFKYM "
    sql = sql & " FROM taItakushaMaster     a,"
    sql = sql & "      tcHogoshaMaster      b,"
    '//基本は振替予定データ
    sql = sql & "      tfFurikaeYoteiData   c "
    sql = sql & " WHERE ABITKB = FAITKB"
    sql = sql & "   AND FAITKB = CAITKB"
    sql = sql & "   AND FAKYCD = CAKYCD"
    'sql = sql & "   AND FAKSCD = CAKSCD"
    sql = sql & "   AND FAHGCD = CAHGCD"
    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN CAFKST AND CAFKED"
    sql = sql & "   AND FASQNO = " & txtFurikaeBi.Number
'//2003/02/03 解約フラグ参照追加
    sql = sql & "   AND NVL(FAKYFG,0) = 0"  '//保護者は解約していない
'//2004/06/03 金額「０」は作成しない
'//2004/06/03 運用が変わる？ので止め！！！
'    sql = sql & "   AND(NVL(faskgk,0) > 0 OR NVL(fahkgk,0) > 0) "
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If dyn.EOF Then
        Call MsgBox(txtFurikaeBi.Text & " に該当するデータはありません.", vbInformation + vbOKOnly, mCaption)
        Exit Sub
    End If
    Dim st As New StructureClass, tmp As String
    Dim reg As New RegistryClass
    Dim mFile As New FileClass, FileName As String, TmpFname As String
    
    dlgFile.DialogTitle = "名前を付けて保存(" & mCaption & ")"
    dlgFile.FileName = reg.OutputFileName(mCaption)
    If IsEmpty(mFile.SaveDialog(dlgFile)) Then
        Exit Sub
    End If
    
    Dim ms As New MouseClass
    Call ms.Start
    
    reg.OutputFileName(mCaption) = dlgFile.FileName
    Call st.SelectStructure(st.KouzaFurikae)

Debug.Print "start= " & Now
    Me.pgbRecord.Visible = True
    Me.pgbRecord.ZOrder 0
    Me.pgbRecord.min = 0
    Me.pgbRecord.Max = dyn.RecordCount
    
    '//取り敢えずテンポラリに書く
    Dim fp As Integer, cnt As Long, SumGaku As Currency
    fp = FreeFile
    TmpFname = mFile.MakeTempFile
    Open TmpFname For Append As #fp
    Do Until dyn.EOF
        DoEvents
        If mAbort Then
            GoTo cmdExport_ClickError
        End If
        tmp = ""
        tmp = tmp & st.SetData(dyn.Fields("ABITCD"), 0)     '委託者番号             '//この項目は委託者マスタ
        tmp = tmp & st.SetData(dyn.Fields("FAKYCD"), 1)     '契約者番号
        tmp = tmp & st.SetData(dyn.Fields("FAHGCD"), 2)     '保護者番号
        tmp = tmp & st.SetData(dyn.Fields("FAKKBN"), 3)     '金融機関区分
        tmp = tmp & st.SetData(dyn.Fields("FABANK"), 4)     '銀行コード
        tmp = tmp & st.SetData(dyn.Fields("FASITN"), 5)     '支店コード
        If eBankKubun.KinnyuuKikan = dyn.Fields("FAKKBN") Then
            tmp = tmp & st.SetData(dyn.Fields("FAKZSB"), 6) '口座種別：預金種目
        Else
            tmp = tmp & st.SetData("0", 6)                  '口座種別：郵便局は「０」
        End If
        tmp = tmp & st.SetData(dyn.Fields("FAKZNO"), 7)     '口座番号
        tmp = tmp & st.SetData(dyn.Fields("FAYBTK"), 8)     '郵便局：通帳記号
        tmp = tmp & st.SetData(dyn.Fields("FAYBTN"), 9)     '郵便局：通帳番号
        tmp = tmp & st.SetData(dyn.Fields("FAKZNM"), 10)    '口座名義人名(カナ)
        tmp = tmp & st.SetData(dyn.Fields("FAFKYM"), 11)    '振替開始年月：FAFKST=>ＳＱＬ編集済み
        tmp = tmp & st.SetData("", 12)                      'filler
        Print #fp, tmp
        cnt = cnt + 1
        Me.pgbRecord.Value = cnt
'////////////////////////////////////////////
'//2012/07/11 スピードアップ改善：ここから
#If cSPEEDUP = False Then
''''//2003/02/03 更新状態フラグ追加:0=DB作成,1=予定作成,2=予定取込,3=請求作成
'''        sql = "UPDATE tfFurikaeYoteiData SET "
'''        sql = sql & " FAUPFG = " & IIf(chkJisseki.Value = eCheckButton.Yotei, _
'''                                        eKouFuriKubun.YoteiText, _
'''                                        eKouFuriKubun.SeikyuText _
'''                                ) & ","
'''        sql = sql & " FAUSID = '" & gdDBS.LoginUserName & "',"
'''        sql = sql & " FAUPDT = SYSDATE"
'''        sql = sql & " WHERE FAITKB = '" & dyn.Fields("FAITKB").Value & "'"
'''        sql = sql & "   AND FAKYCD = '" & dyn.Fields("FAKYCD").Value & "'"
'''        'sql = sql & "   AND FAKSCD = '" & dyn.Fields("FAKSCD").Value & "'"
'''        sql = sql & "   AND FAHGCD = '" & dyn.Fields("FAHGCD").Value & "'"
'''        sql = sql & "   AND FASQNO = " & txtFurikaeBi.Number
'''        Call gdDBS.Database.ExecuteSQL(sql)
#End If
'//2012/07/11 スピードアップ改善：ここまで
'////////////////////////////////////////////
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    Me.pgbRecord.Visible = False
    lblMessage.ZOrder 0
    Me.Refresh
'////////////////////////////////////////////
'//2012/07/11 スピードアップ改善：ここから
#If cSPEEDUP = True Then
    sql = "UPDATE tfFurikaeYoteiData SET "
    sql = sql & " FAUPFG = " & IIf(chkJisseki.Value = eCheckButton.Yotei, _
                                    eKouFuriKubun.YoteiText, _
                                    eKouFuriKubun.SeikyuText _
                            ) & ","
    sql = sql & " FAUSID = '" & gdDBS.LoginUserName & "',"
    sql = sql & " FAUPDT = SYSDATE"
    sql = sql & " WHERE FASQNO = " & txtFurikaeBi.Number
    Dim updCnt As Long
    updCnt = gdDBS.Database.ExecuteSQL(sql)
    If updCnt <> cnt Then
        Call err.Raise(-1, "cmdMakeText", "テキスト作成は失敗しました." & vbCrLf & "ＤＢ作成後に各マスタが変更された可能性があります.")
    End If
#End If
'//2012/07/11 スピードアップ改善：ここまで
'////////////////////////////////////////////

Debug.Print "  end= " & Now
    
    Close #fp
#If 1 Then
    '//ファイル移動     MOVEFILE_REPLACE_EXISTING=Replace , MOVEFILE_COPY_ALLOWED=Copy & Delete
    Call MoveFileEx(TmpFname, reg.OutputFileName(mCaption), MOVEFILE_REPLACE_EXISTING + MOVEFILE_COPY_ALLOWED)
    'Call MoveFileEx(TmpFname, reg.FileName(mCaption), MOVEFILE_REPLACE_EXISTING)
#Else
    '//ファイルコピー
    Call FileCopy(TmpFname, reg.FileName(mCaption))
#End If
    Set mFile = Nothing
    '//実行更新フラグ設定：この関数は予定・請求ともに実行可能
    Select Case chkJisseki.Value
    Case eCheckButton.Yotei
        gdDBS.SystemUpdate("AAUPD2") = 1
    Case eCheckButton.Kakutei
        gdDBS.SystemUpdate("AAUPD3") = 1
    End Select
    Call gdDBS.AutoLogOut(mCaption, "テキスト作成(" & txtFurikaeBi & " : " & cnt & " 件)")
'//2004/04/26 新規件数&０円件数＆合計金額の追加
'//2004/05/17 詳細を削除
    lblMessage.Caption = cnt & " 件のデータが作成されました。"
                    '// & vbCrLf & _
                        "<< 詳細 >>" & vbCrLf & _
                        "新規件数 = " & NewCnt & vbCrLf & _
                        "  ０円件数 = " & ZeroCnt & vbCrLf & _
                        "合計金額 = " & Format(TotalGaku, "#,##0")
    Exit Sub
cmdExport_ClickError:
    Call gdDBS.ErrorCheck       '//エラートラップ
    Set mFile = Nothing
End Sub

Private Sub cmdSend_Click()
    Dim reg As New RegistryClass
    Call Shell(reg.TransferCommand(mCaption), vbNormalFocus)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    Call mForm.LockedControl(False)
    lblMessage.Caption = mExeMsg
    'txtFurikaeBi.Number = gdDBS.SYSDATE("YYYYMMDD")
    txtFurikaeBi.Number = gdDBS.Nz(gdDBS.SystemUpdate("AANXKZ"))
    chkJisseki.Value = eCheckButton.Mukou  '無効に設定
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call gdDBS.Database.Rollback
    mAbort = True
    Set frmKouzaFurikaeExport = Nothing
    Set mForm = Nothing
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

Private Sub txtFurikaeBi_DropOpen(NoDefault As Boolean)
    txtFurikaeBi.Calendar.Holidays = gdDBS.Holiday(txtFurikaeBi.Year)
End Sub

