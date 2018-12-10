VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
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
   Begin VB.Frame fraColors 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   6360
      TabIndex        =   10
      Top             =   -90
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
      Left            =   5040
      TabIndex        =   8
      Top             =   -90
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
      Left            =   3720
      TabIndex        =   6
      Top             =   -90
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
      Left            =   1020
      TabIndex        =   0
      Top             =   120
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   "保護者マスタ履歴照会.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタ履歴照会.frx":006E
      Key             =   "保護者マスタ履歴照会.frx":008C
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
      MaxLength       =   5
      LengthAsByte    =   0
      Text            =   "12345"
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
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
      Left            =   1980
      TabIndex        =   1
      Top             =   60
      Width           =   1455
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
      Bindings        =   "保護者マスタ履歴照会.frx":00D0
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
      SpreadDesigner  =   "保護者マスタ履歴照会.frx":00F2
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
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "契約者"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   180
      Width           =   675
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   255
      Left            =   11100
      TabIndex        =   4
      Top             =   60
      Width           =   1395
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

Private Enum eRecord
    eRireki = 0
    eMaster = 1
    eDefaultColor = 0
    eKaiyakuColor
    eRirekiColor
End Enum

Private Enum eSprCol
    eRireki = 1
    eKaiyaku = 21
End Enum

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
    FieldNames = Array("R区分", "保護者", "教室", "ＳＥＱ", "保護者名", "口座名義人", "生徒氏名", _
                       "区分", "銀行名", "支店名", "種別", "口座番号", "記号", "通帳番号", _
                       "契約開始日", "契約終了日", "振替開始日", "振替終了日", "請求額", "変更額", _
                       "解約", _
                       "解約処理日", _
                       "新規扱い日", _
                       "リスト出力日", "更新者ＩＤ", _
                       "作成日", "更新日" _
                )
    '////////////////////////
    '//表示する項目の編集
    FieldIDs = Array("rKUBUN", "CAHGCD", "CAKSCD", "to_char(to_date(CASQNO,'yyyymmdd'),'yyyy/mm/dd')", "CAKJNM", "CAKZNM", "CASTNM", _
                     "cakkbnX", "dabknm", "dastnm", "cakzsbX", "CAKZNO", "CAYBTK", "CAYBTN", _
                     "to_char(to_date(decode(CAKYST,0,null,CAKYST),'yyyymmdd'),'yyyy/mm/dd')", _
                     "to_char(to_date(decode(CAKYED,0,null,CAKYED),'yyyymmdd'),'yyyy/mm/dd')", _
                     "to_char(to_date(decode(CAFKST,0,null,CAFKST),'yyyymmdd'),'yyyy/mm/dd')", _
                     "to_char(to_date(decode(CAFKED,0,null,CAFKED),'yyyymmdd'),'yyyy/mm/dd')", _
                     "to_char(CASKGK,'999,999,999')", "to_char(CAHKGK,'999,999,999')", _
                     "cakyfgX", _
                     "to_char(CAKYSR,'yyyy/mm/dd hh24:mi:ss')", _
                     "to_char(CANWDT,'yyyy/mm/dd hh24:mi:ss')", _
                     "to_char(CACHEK,'yyyy/mm/dd hh24:mi:ss')", "CAUSID", _
                     "to_char(CAADDT,'yyyy/mm/dd hh24:mi:ss')", _
                     "to_char(CAUPDT,'yyyy/mm/dd hh24:mi:ss')" _
                )
    ReDim ColWidths(UBound(FieldNames))
    '////////////////////////
    '//表示する列幅
    'defualt = 8.0
    ColWidths = Array(0, 5.1, 4, 9.5, 14, 14, 14, 6, 12, 12, 4, 7, 3.5, 7.5, 9.5, 9.5, 9.5, 9.5, _
                      8, 8, 4, 16, 16, 16, 16, 16, 16)
    sprRireki.Row = -1  '//全行が対象
    sprRireki.MaxCols = UBound(FieldIDs) + 1
    sprRireki.ColsFrozen = 3
    For ix = LBound(FieldIDs) To UBound(FieldIDs)
        mSpread.ColWidth(ix + 1) = ColWidths(ix)
        sprRireki.Col = ix + 1      '//指定列をフォーマット
        Select Case FieldNames(ix)
        Case "保護者名", "生徒氏名", "金融機関", "銀行名", "支店名", "口座番号", "口座名義人", "通帳番号", "更新者ＩＤ"
            sprRireki.TypeHAlign = TypeHAlignLeft
        Case "請求額", "変更額"
            sprRireki.TypeHAlign = TypeHAlignRight
        Case Else
            sprRireki.TypeHAlign = TypeHAlignCenter
        End Select
    Next ix
    '////////////////////////
    '//ＤＢ取得項目
    IDs = Array("CAKYCD", "CAKSCD", "CAHGCD", "CASQNO", "CAKJNM", "CAKNNM", "CASTNM", _
                "cakkbn", "cabank", "casitn", "cakzsb", "CAKZNO", "CAKZNM", "CAYBTK", "CAYBTN", _
                "CAKYST", "CAKYED", "CAFKST", "CAFKED", "CASKGK", "CAHKGK", _
                "CAKYSR", "cakyfg", "CANWDT", "CACHEK", "CAUSID", "CAADDT", "CAUPDT")
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
    
    sql = "SELECT " & vbCrLf
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
        sql = sql & " decode(b.DAKJNM,null,CABANK,b.DAKJNM) DABKNM," & vbCrLf
        sql = sql & " decode(c.DAKJNM,null,CASITN,c.DAKJNM) DASTNM " & vbCrLf
        sql = sql & " FROM tcHogoshaMaster  a," & vbCrLf
        sql = sql & "      tdBankMaster     b," & vbCrLf
        sql = sql & "      tdBankMaster     c," & vbCrLf
        sql = sql & "      taItakushaMaster d " & vbCrLf
        sql = sql & " WHERE CABANK = b.DABANK(+)" & vbCrLf
        sql = sql & "   AND '000'  = b.DASITN(+)" & vbCrLf
        sql = sql & "   AND ':'    = b.DASQNO(+)" & vbCrLf
        sql = sql & "   AND CABANK = c.DABANK(+)" & vbCrLf
        sql = sql & "   AND CASITN = c.DASITN(+)" & vbCrLf
        sql = sql & "   AND 'ｱ'    = c.DASQNO(+)" & vbCrLf
        sql = sql & "   AND CAITKB = ABITKB " & vbCrLf
        If "" = Trim(txtCAKYCD.Text) Then
            sql = sql & " AND CAKYCD IS NULL"
        Else
            sql = sql & " AND CAKYCD = " & gdDBS.ColumnDataSet(txtCAKYCD.Text, vEnd:=True) & vbCrLf
        End If
        sql = sql & " UNION ALL " & vbCrLf
        '///////////////////////////////
        '//保護者履歴の内容
        '///////////////////////////////
        sql = sql & "SELECT " & vbCrLf
        For ix = LBound(IDs) To UBound(IDs)
            Select Case UCase(IDs(ix))
            Case UCase("CAKYSR"), UCase("CANWDT")
                sql = sql & " null " & IDs(ix) & ","
            Case Else
                sql = sql & IDs(ix) & ","
            End Select
        Next ix
        sql = sql & " 0 rKUBUN,CAMKDT," & vbCrLf
        sql = sql & " DECODE(CAKKBN,0,NULL,1,'郵便局',NULL) CAKKBNx," & vbCrLf
        sql = sql & " DECODE(CAKKBN,0,DECODE(CAKZSB,1,'普通',2,'当座',NULL),NULL) CAKZSBx," & vbCrLf
        sql = sql & " DECODE(CAKYFG,0,NULL,1,'解約',NULL) CAKYFGx," & vbCrLf
        sql = sql & " decode(b.DAKJNM,null,CABANK,b.DAKJNM) DABKNM," & vbCrLf
        sql = sql & " decode(c.DAKJNM,null,CASITN,c.DAKJNM) DASTNM " & vbCrLf
        sql = sql & " FROM tcHogoshaMasterRireki  a," & vbCrLf
        sql = sql & "      tdBankMaster     b," & vbCrLf
        sql = sql & "      tdBankMaster     c," & vbCrLf
        sql = sql & "      taItakushaMaster d " & vbCrLf
        sql = sql & " WHERE CABANK = b.DABANK(+)" & vbCrLf
        sql = sql & "   AND '000'  = b.DASITN(+)" & vbCrLf
        sql = sql & "   AND ':'    = b.DASQNO(+)" & vbCrLf
        sql = sql & "   AND CABANK = c.DABANK(+)" & vbCrLf
        sql = sql & "   AND CASITN = c.DASITN(+)" & vbCrLf
        sql = sql & "   AND 'ｱ'    = c.DASQNO(+)" & vbCrLf
        sql = sql & "   AND CAITKB = ABITKB " & vbCrLf
        If "" = Trim(txtCAKYCD.Text) Then
            sql = sql & " AND CAKYCD IS NULL"
        Else
            sql = sql & " AND CAKYCD = " & gdDBS.ColumnDataSet(txtCAKYCD.Text, vEnd:=True) & vbCrLf
        End If
    sql = sql & ")" & vbCrLf
    sql = sql & " ORDER BY CAKYCD,CAHGCD,cakscd,rkubun desc,CASQNO desc,CAMKDT DESC" & vbCrLf
    dbcHogoshaMstRireki.RecordSource = "select * from (" & sql & ")"
    dbcHogoshaMstRireki.Refresh
    '//仮想最大行を設定しなおししないとデータが正常に表示されない
    '//2012/07/02 特定のデータに対して表示ができない？バグ？なので設定行をコメント化：SQLが悪かった？SQLに "select * from (" & sql & ")" で回避
    sprRireki.VirtualMaxRows = dbcHogoshaMstRireki.Recordset.RecordCount
    sprRireki.VisibleRows = sprRireki.VirtualMaxRows
    sprRireki.VirtualMode = True
    'sprRireki.OperationMode = OperationModeRow
    cmdSearch.Enabled = True
    Call sprRireki_TopLeftChange(1, 1, 1, 1)    '//履歴行の行カラー変更を強制する
cmdSearch_ClickError:
    cmdSearch.Enabled = True
End Sub

Private Sub Form_Load()
    Call mForm.Init(Me, gdDBS)
    Call mSpread.Init(sprRireki)
    txtCAKYCD.Text = "" '"20013"
    'sprRireki.MaxRows = 0
    Call cmdSearch_Click
'    fraColors(eRecord.eDefaultColor).BackColor = RGB(255, 255, 255)
'    fraColors(eRecord.eKaiyakuColor).BackColor = RGB(255, 127, 191)
'    fraColors(eRecord.eRirekiColor).BackColor = RGB(192, 255, 239)
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
