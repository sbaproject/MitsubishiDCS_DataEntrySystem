VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmHogoshaMasterRireki 
   Caption         =   "保護者マスタ履歴 照会"
   ClientHeight    =   7650
   ClientLeft      =   2430
   ClientTop       =   2970
   ClientWidth     =   12750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   12750
   Begin VB.ComboBox cboFurikae 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'ｵﾌ固定
      ItemData        =   "保護者マスタ履歴照会.frx":0000
      Left            =   2760
      List            =   "保護者マスタ履歴照会.frx":0010
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   60
      Left            =   7140
      TabIndex        =   12
      Top             =   0
      Width           =   3975
   End
   Begin VB.Frame fraColors 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   9840
      TabIndex        =   10
      Top             =   -30
      Width           =   1215
      Begin VB.Label lblColors 
         Alignment       =   2  '中央揃え
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "履　歴"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   11
         Top             =   180
         Width           =   585
      End
   End
   Begin VB.Frame fraColors 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   8520
      TabIndex        =   8
      Top             =   -30
      Width           =   1215
      Begin VB.Label lblColors 
         Alignment       =   2  '中央揃え
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "解　約"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   9
         Top             =   180
         Width           =   585
      End
   End
   Begin VB.Frame fraColors 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   7200
      TabIndex        =   6
      Top             =   -30
      Width           =   1215
      Begin VB.Label lblColors 
         Alignment       =   2  '中央揃え
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "通　常"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   7
         Top             =   180
         Width           =   585
      End
   End
   Begin imText6Ctl.imText txtCAKYCD 
      Height          =   315
      Left            =   900
      TabIndex        =   0
      Top             =   120
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "保護者マスタ履歴照会.frx":0034
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタ履歴照会.frx":00A2
      Key             =   "保護者マスタ履歴照会.frx":00C0
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   "9"
      FormatMode      =   0
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   7
      LengthAsByte    =   0
      Text            =   "1234567"
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   3
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "対象者検索(&S)"
      Height          =   435
      Left            =   4980
      TabIndex        =   1
      Top             =   60
      Width           =   1300
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "終了(&X)"
      Height          =   435
      Left            =   11100
      TabIndex        =   3
      Top             =   7020
      Width           =   1395
   End
   Begin FPSpread.vaSpread sprRireki 
      Bindings        =   "保護者マスタ履歴照会.frx":0104
      Height          =   6255
      Left            =   180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   540
      Width           =   12435
      _Version        =   196608
      _ExtentX        =   21934
      _ExtentY        =   11033
      _StockProps     =   64
      ColsFrozen      =   1
      DAutoCellTypes  =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   1
      MaxRows         =   1
      OperationMode   =   1
      SpreadDesigner  =   "保護者マスタ履歴照会.frx":0126
      UserResize      =   1
      VirtualMode     =   -1  'True
      VisibleCols     =   1
   End
   Begin ORADCLibCtl.ORADC dbcHogoshaMstRireki 
      Height          =   315
      Left            =   6480
      Top             =   7140
      Visible         =   0   'False
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   207
      BackColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   "kumon"
      Connect         =   "kumon/kumon"
      RecordSource    =   "select * from tcHogoshaMasterRireki"
   End
   Begin imDate6Ctl.imDate txtKijunBi 
      Height          =   315
      Left            =   3960
      TabIndex        =   15
      Top             =   120
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1499
      _ExtentY        =   556
      Calendar        =   "保護者マスタ履歴照会.frx":03C9
      Caption         =   "保護者マスタ履歴照会.frx":0549
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタ履歴照会.frx":05B7
      Keys            =   "保護者マスタ履歴照会.frx":05D5
      Spin            =   "保護者マスタ履歴照会.frx":0633
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
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
      MaxDate         =   2958465
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
      Text            =   "2012/12/11"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   41254
      CenturyMode     =   0
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "振替方法"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1935
      TabIndex        =   14
      Top             =   165
      Width           =   780
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   255
      Left            =   11220
      TabIndex        =   4
      Top             =   60
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "ｵｰﾅｰ№"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   165
      Width           =   675
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
Attribute VB_Name = "frmHogoshaMasterRireki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mForm   As New FormClass
Private mSpread As New SpreadClass

'//2014/06/27 履歴 <==> 保護者メンテに飛ぶのでフォーム内容を退避、復活用に宣言
Private mRetForm As Form

Private Enum eFurikae
    eALL
    ePaper
    eBank
    eKaiyaku
End Enum

Private Enum eRecord
    eRireki = 0
    eMaster = 1
    eDefaultColor = 0
    eKaiyakuColor
    eRirekiColor
End Enum

Private Enum eSprCol
    eRireki = 1
    eCAHGCD = 2
    eKaiyaku = 16
    eCAKYCD = 21
End Enum

Private Sub cboFurikae_Click()
    txtKijunBi.Visible = eFurikae.ePaper = cboFurikae.ListIndex Or eFurikae.eBank = cboFurikae.ListIndex
    txtKijunBi.Value = Now()
End Sub

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Dim FieldNames As Variant
    Dim FieldIDs As Variant, IDs As Variant
    Dim ColWidths As Variant
    Dim ix As Integer
    Dim ms As New MouseClass
    
    cmdSearch.Enabled = False
    Call ms.Start
    '////////////////////////
    '//表示する名前
    FieldNames = Array("R区分", "保護者", "ＳＥＱ", "保護者名", "口座名義人", "生徒氏名", _
                       "金融機関", "銀行名", "支店名", "種別", "口座番号", "記号", "通帳番号", _
                       "振替開始", "振替終了", _
                       "解約", _
                       "新規扱い日", _
                       "更新者", _
                       "データ作成日", "データ更新日", "ｵｰﾅｰNo" _
                )
    '////////////////////////
    '//表示する項目の編集
    '2012/11/15 CASQNO に －１ があるので ==> (case when length(CASQNO)=8 then casqno else null end)
    FieldIDs = Array("rKUBUN", "CAHGCD", _
            "to_char(to_date((case when length(CASQNO)=8 then casqno else null end),'yyyymmdd'),'yyyy/mm/dd')", _
                     "CAKJNM", "CAKZNM", "CASTNM", _
                     "cakkbnX", "dabknm", "dastnm", "cakzsbX", "CAKZNO", "CAYBTK", "CAYBTN", _
                     "to_char(to_date(decode(CAFKST,0,null,CAFKST),'yyyymmdd'),'yyyy/mm')", _
                     "to_char(to_date(decode(CAFKED,0,null,CAFKED),'yyyymmdd'),'yyyy/mm')", _
                     "cakyfgX", _
                     "to_char(CANWDT,'yyyy/mm/dd hh24:mi:ss')", _
                     "CAUSID", _
                     "to_char(CAADDT,'yyyy/mm/dd hh24:mi:ss')", _
                     "to_char(CAUPDT,'yyyy/mm/dd hh24:mi:ss')", "cakycd" _
                )
    ReDim ColWidths(UBound(FieldNames))
    '////////////////////////
    '//表示する列幅
    'defualt = 8.0
    ColWidths = Array(0, 7.6, 9.5, 14, 14, 14, 6, 12, 12, 4, 7, 3.5, 7.5, 8, 8, _
                      4, 16, 10, 16, 16, 7.6)
    sprRireki.Row = -1  '//全行が対象
    sprRireki.MaxCols = UBound(FieldIDs) + 1
    sprRireki.ColsFrozen = 3
    For ix = LBound(FieldIDs) To UBound(FieldIDs)
        mSpread.ColWidth(ix + 1) = ColWidths(ix)
        sprRireki.Col = ix + 1      '//指定列をフォーマット
        Select Case FieldNames(ix)
        Case "保護者名", "生徒氏名", "金融機関", "銀行名", "支店名", "口座番号", "口座名義人", "通帳番号", "更新者"
            sprRireki.TypeHAlign = TypeHAlignLeft
        Case Else
            sprRireki.TypeHAlign = TypeHAlignCenter
        End Select
    Next ix
    '////////////////////////
    '//ＤＢ取得項目
    IDs = Array("CAKYCD", "CAHGCD", "CASQNO", "CAKJNM", "CAKNNM", "CASTNM", _
                "cakkbn", "cabank", "casitn", "cakzsb", "CAKZNO", "CAKZNM", "CAYBTK", "CAYBTN", _
                "CAFKST", "CAFKED", _
                "cakyfg", "CANWDT", "CAUSID", "CAADDT", "CAUPDT")
    Dim sql As String
    
    On Error GoTo cmdSearch_ClickError
'    sql = "SELECT * "
'    For ix = LBound(mFieldNames) To UBound(mFieldNames)
'        sql = sql & IDs(ix) & " " & mFieldNames(ix) & ","
'    Next ix
'    sql = Left(sql, Len(sql) - 1)
'    sql = sql & " FROM tcHogoshaMasterRireki "
'    If "" <> Trim(txtCAKYCD.Text) Then
'        sql = sql & " WHERE CAKYCD = " & gdDBS.ColumnDataSet(txtCAKYCD.Text, vEnd:=True)
'    End If

    sql = "with vdBankMaster as("
    sql = sql & " select"
    sql = sql & " a.darkbn,a.dabank,a.daknnm,a.dakjnm,b.dasitn,b.daknnm dastkn,b.dakjnm dastkj,b.dasqno,b.dahtif"
    sql = sql & " from TDBANKMASTER a,TDBANKMASTER b"
    sql = sql & " Where a.dabank = b.dabank"
    sql = sql & "   and a.dasqno=':'"
    sql = sql & "   and b.dasqno='ｱ'"  '--"ｱ"以外は無い
    sql = sql & " order by a.dabank,b.dasitn"
    sql = sql & ")," & vbCrLf
    sql = sql & " vcHogoshaMaster as("
    sql = sql & " select a.* from tcHogoshaMaster a"
    sql = sql & " where (caitkb,cakycd,cahgcd,casqno) in("
    sql = sql & "       select caitkb,cakycd,cahgcd,max(casqno)"
    sql = sql & "       from tcHogoshaMaster "
    sql = sql & "       group by caitkb,cakycd,cahgcd"
    sql = sql & "   )"
    sql = sql & ")" & vbCrLf
    
    sql = sql & "SELECT " & vbCrLf
    For ix = LBound(FieldIDs) To UBound(FieldIDs)
        sql = sql & FieldIDs(ix) & " " & FieldNames(ix) & ","
    Next ix
    sql = Left(sql, Len(sql) - 1)
    sql = sql & " FROM(" & vbCrLf
        '///////////////////////////////
        '//保護者マスターの内容
        '///////////////////////////////
        sql = sql & "SELECT " & vbCrLf
        For ix = LBound(IDs) To UBound(IDs)
            sql = sql & IDs(ix) & ","
        Next ix
        sql = sql & " 1 rKUBUN,SYSDATE CAMKDT," & vbCrLf
        sql = sql & " DECODE(CAKKBN,0,NULL,1,'郵便局','その他') CAKKBNx," & vbCrLf
        sql = sql & " DECODE(CAKKBN,0,DECODE(CAKZSB,1,'普通',2,'当座',NULL),NULL) CAKZSBx," & vbCrLf
        sql = sql & " DECODE(CAKYFG,0,NULL,1,'解約','其他') CAKYFGx," & vbCrLf
        sql = sql & " decode(b.DAKJNM,null,CABANK, b.DAKJNM) DABKNM," & vbCrLf
        sql = sql & " decode(b.DASTKJ,null,CASITN, b.DASTKJ) DASTNM " & vbCrLf
'//2015/02/09 保護者マスタの本体の口座変更した(レコード追加)場合変更前が出ないので変更
       'sql = sql & " FROM vcHogoshaMaster  a," & vbCrLf
        sql = sql & " FROM tcHogoshaMaster  a," & vbCrLf
        sql = sql & "      vdBankMaster     b," & vbCrLf
        sql = sql & "      taItakushaMaster d " & vbCrLf
        sql = sql & " WHERE CABANK = b.DABANK(+)" & vbCrLf
        sql = sql & "   AND CASITN = b.DASITN(+)" & vbCrLf
        sql = sql & "   AND CAITKB = ABITKB " & vbCrLf
        If "" <> Trim(txtCAKYCD.Text) Then
'//2015/02/09 LIKE 文に変更
           'sql = sql & " AND CAKYCD = " & gdDBS.ColumnDataSet(txtCAKYCD.Text, vEnd:=True) & vbCrLf
            sql = sql & " AND CAKYCD LIKE " & gdDBS.ColumnDataSet("%" & txtCAKYCD.Text & "%", vEnd:=True) & vbCrLf
        End If
        Select Case cboFurikae.ListIndex
        Case eFurikae.eALL
        Case eFurikae.ePaper
            sql = sql & " and cafkst > " & Left(txtKijunBi.Number, 6) & "01" & vbCrLf
            sql = sql & " and nvl(cakyfg,'0') = '0' " & vbCrLf
        Case eFurikae.eBank
            sql = sql & " and " & Left(txtKijunBi.Number, 6) & "01" & " between cafkst and cafked " & vbCrLf
            sql = sql & " and nvl(cakyfg,'0') = '0' " & vbCrLf
        Case eFurikae.eKaiyaku
            sql = sql & " and nvl(cakyfg,'0') <> '0' " & vbCrLf
        End Select
        sql = sql & " UNION ALL " & vbCrLf
        '///////////////////////////////
        '//保護者履歴の内容
        '///////////////////////////////
        sql = sql & "SELECT " & vbCrLf
        For ix = LBound(IDs) To UBound(IDs)
            Select Case UCase(IDs(ix))
            Case UCase("CANWDT")
                sql = sql & " null " & IDs(ix) & ","
            Case Else
                sql = sql & IDs(ix) & ","
            End Select
        Next ix
        sql = sql & " 0 rKUBUN,CAMKDT," & vbCrLf
        sql = sql & " DECODE(CAKKBN,0,NULL,1,'郵便局',NULL) CAKKBNx," & vbCrLf
        sql = sql & " DECODE(CAKKBN,0,DECODE(CAKZSB,1,'普通',2,'当座',NULL),NULL) CAKZSBx," & vbCrLf
        sql = sql & " DECODE(CAKYFG,0,NULL,1,'解約',NULL) CAKYFGx," & vbCrLf
        sql = sql & " decode(b.DAKJNM,null,CABANK, b.DAKJNM) DABKNM," & vbCrLf
        sql = sql & " decode(b.DASTKJ,null,CASITN, b.DASTKJ) DASTNM " & vbCrLf
        sql = sql & " FROM tcHogoshaMasterRireki  a," & vbCrLf
        sql = sql & "      vdBankMaster     b," & vbCrLf
        sql = sql & "      taItakushaMaster d " & vbCrLf
        sql = sql & " WHERE CABANK = b.DABANK(+)" & vbCrLf
        sql = sql & "   AND CASITN = b.DASITN(+)" & vbCrLf
        sql = sql & "   AND CAITKB = ABITKB " & vbCrLf
        If "" <> Trim(txtCAKYCD.Text) Then
'//2015/02/09 LIKE 文に変更
           'sql = sql & " AND CAKYCD = " & gdDBS.ColumnDataSet(txtCAKYCD.Text, vEnd:=True) & vbCrLf
            sql = sql & " AND CAKYCD LIKE " & gdDBS.ColumnDataSet("%" & txtCAKYCD.Text & "%", vEnd:=True) & vbCrLf
        End If
        If eFurikae.eALL < cboFurikae.ListIndex Then
            sql = sql & "   AND(CAKYCD,CAHGCD) in( "
            sql = sql & "   select CAKYCD,CAHGCD"
            sql = sql & "   FROM vcHogoshaMaster  a," & vbCrLf
            sql = sql & "        vdBankMaster     b," & vbCrLf
            sql = sql & "        taItakushaMaster d " & vbCrLf
            sql = sql & "   WHERE CABANK = b.DABANK(+)" & vbCrLf
            sql = sql & "     AND CASITN = b.DASITN(+)" & vbCrLf
            sql = sql & "     AND CAITKB = ABITKB " & vbCrLf
            If "" <> Trim(txtCAKYCD.Text) Then
                sql = sql & " AND CAKYCD = " & gdDBS.ColumnDataSet(txtCAKYCD.Text, vEnd:=True) & vbCrLf
            End If
            Select Case cboFurikae.ListIndex
            Case eFurikae.eALL
            Case eFurikae.ePaper
                sql = sql & " and cafkst > " & Left(txtKijunBi.Number, 6) & "01" & vbCrLf
                sql = sql & " and nvl(cakyfg,'0') = '0' " & vbCrLf
            Case eFurikae.eBank
                sql = sql & " and " & Left(txtKijunBi.Number, 6) & "01" & " between cafkst and cafked " & vbCrLf
                sql = sql & " and nvl(cakyfg,'0') = '0' " & vbCrLf
            Case eFurikae.eKaiyaku
                sql = sql & " and nvl(cakyfg,'0') <> '0' " & vbCrLf
            End Select
            sql = sql & ")" & vbCrLf
        End If
    sql = sql & ")" & vbCrLf
    'sql = sql & " ORDER BY CAKYCD,CAHGCD,CASQNO,CAMKDT DESC" & vbCrLf
    sql = sql & " ORDER BY CAKYCD,CAHGCD,CASQNO desc,rkubun desc,CAMKDT DESC" & vbCrLf
    dbcHogoshaMstRireki.RecordSource = "select * from(" & sql & ")"
    dbcHogoshaMstRireki.Refresh
    '//仮想最大行を設定しなおししないとデータが正常に表示されない
    sprRireki.VirtualMaxRows = dbcHogoshaMstRireki.Recordset.RecordCount
    sprRireki.VisibleRows = sprRireki.VirtualMaxRows
    sprRireki.VirtualMode = True
    'sprRireki.OperationMode = OperationModeRow
    cmdSearch.Enabled = True
    Call sprRireki_TopLeftChange(1, 1, 1, 1)    '//履歴行の行カラー変更を強制する
cmdSearch_ClickError:
    cmdSearch.Enabled = True
End Sub

Private Sub Form_Activate()
    '//2014/06/27 保護者マスタを強制破棄
    Unload frmHogoshaMaster
End Sub

Private Sub Form_Load()
    '//2014/06/27 履歴へ飛ぶのでメニューを退避
    Set mRetForm = gdForm
    Call mForm.Init(Me, gdDBS)
    Call mSpread.Init(sprRireki)
    cboFurikae.Clear
    Call cboFurikae.AddItem("全て", eFurikae.eALL)
    Call cboFurikae.AddItem("振替用紙", eFurikae.ePaper)
    Call cboFurikae.AddItem("口座振替", eFurikae.eBank)
    Call cboFurikae.AddItem("解約", eFurikae.eKaiyaku)
    'cboFurikae.ItemData(eFurikae.eALL) = eFurikae.eALL
    'cboFurikae.ItemData(eFurikae.ePaper) = eFurikae.ePaper
    'cboFurikae.ItemData(eFurikae.eBank) = eFurikae.eBank
    'cboFurikae.ItemData(eFurikae.eKaiyaku) = eFurikae.eKaiyaku
    cboFurikae.ListIndex = eFurikae.eALL
    
    '//列を表示する為にブランクを設定して検索をする＝０件表示
    txtCAKYCD.Text = " " '"20013"
    Call cmdSearch_Click
    txtCAKYCD.Text = ""
    sprRireki.MaxRows = 0
'    fraColors(eRecord.eDefaultColor).BackColor = RGB(255, 255, 255)
'    fraColors(eRecord.eKaiyakuColor).BackColor = RGB(255, 127, 191)
'    fraColors(eRecord.eRirekiColor).BackColor = RGB(192, 255, 239)
    Call cboFurikae_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmHogoshaMasterRireki = Nothing
    Set mForm = Nothing
    '//2014/06/27 履歴から保護者メンテに飛ぶのでメニューに復活
    Set gdForm = mRetForm
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

Private Sub sprRireki_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Col < 1 And Row < 1 Then
        Exit Sub
    End If
    Dim frm As Form
    Set frm = frmHogoshaMaster
    Call frm.Show
    frm.txtCAKYCD.Text = mSpread.Text(eSprCol.eCAKYCD, Row)
    frm.txtCAHGCD.Text = mSpread.Text(eSprCol.eCAHGCD, Row)
    Call frm.txtCAHGCD_KeyDown(vbKeyReturn, 0)
    Set gdForm = Me
End Sub

Private Sub sprRireki_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    Dim Row As Long, data As Variant
    'sprRireki.BlockMode = True
    For Row = NewTop To NewTop + 24
        If Row <= mSpread.MaxRows Then
            mSpread.BackColor(-1, Row) = fraColors(eRecord.eDefaultColor).BackColor
            '//履歴情報？
            If eRecord.eMaster <> mSpread.Text(eSprCol.eRireki, Row) Then
                mSpread.BackColor(-1, Row) = fraColors(eRecord.eRirekiColor).BackColor
            Else
                '//解約状態？
                If "" <> mSpread.Text(eSprCol.eKaiyaku, Row) Then
                    mSpread.BackColor(-1, Row) = fraColors(eRecord.eKaiyakuColor).BackColor
                End If
            End If
        End If
    Next Row
    'sprRireki.BlockMode = False
End Sub

Private Sub txtCAKYCD_KeyDown(KeyCode As Integer, Shift As Integer)
    '// Return または Shift＋TAB のときのみ処理する
    If Not (KeyCode = vbKeyReturn) Then
        Exit Sub
    ElseIf 0 = Len(Trim(txtCAKYCD.Text)) Then
        Exit Sub
    End If
'//2013/06/18 前ゼロ埋め込み
    txtCAKYCD.Text = Format(Val(txtCAKYCD.Text), String(7, "0"))
End Sub
