VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Begin VB.Form frmYoteiReqImportReport 
   Caption         =   "振替予定表取込 結果(印刷)"
   ClientHeight    =   4065
   ClientLeft      =   3795
   ClientTop       =   2355
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6360
   Begin VB.ComboBox cboImpDate 
      Height          =   300
      ItemData        =   "振替予定表取込_結果(印刷).frx":0000
      Left            =   1920
      List            =   "振替予定表取込_結果(印刷).frx":000D
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   1320
      Width           =   2115
   End
   Begin VB.CheckBox chkDefault 
      Caption         =   "前回累積日"
      Height          =   315
      Left            =   3900
      TabIndex        =   6
      Top             =   420
      Width           =   1575
   End
   Begin imText6Ctl.imText txtStartDate 
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   420
      Width           =   1875
      _Version        =   65536
      _ExtentX        =   3307
      _ExtentY        =   556
      Caption         =   "振替予定表取込_結果(印刷).frx":0055
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替予定表取込_結果(印刷).frx":00C3
      Key             =   "振替予定表取込_結果(印刷).frx":00E1
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
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
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   "2004/06/28 12:13:14"
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "印刷(&P)"
      Height          =   435
      Left            =   480
      TabIndex        =   1
      ToolTipText     =   "印刷を開始する場合"
      Top             =   3300
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "終了(&E)"
      Height          =   435
      Left            =   4680
      TabIndex        =   2
      ToolTipText     =   "この作業を終了してメインメニューに戻る場合"
      Top             =   3300
      Width           =   1335
   End
   Begin VB.Label lblImpDate 
      Alignment       =   1  '右揃え
      Caption         =   "取込日時"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   1380
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "基準日"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label1"
      Height          =   195
      Left            =   4860
      TabIndex        =   3
      Top             =   0
      Width           =   1275
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
Attribute VB_Name = "frmYoteiReqImportReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm As New FormClass
Private mCaption As String
Private mStartDate As String

Private Sub chkDefault_Click()
    If 0 = chkDefault.Value Then
        txtStartDate.Enabled = True
    Else
        txtStartDate.Text = mStartDate
        txtStartDate.Enabled = False
    End If
End Sub

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim ms As New MouseClass
    Call ms.Start
    Dim sql As String
    
    sql = "SELECT "
    sql = sql & " FIADID,"                                   '//処理結果
    sql = sql & " ABITCD,ABKJNM,FIKYCD,"
    sql = sql & " MAX(FIMKDT)   FIMKDT,"
    sql = sql & " SUM(ALLCNT  ) ALLCNT  ," '//件数
    sql = sql & " SUM(FIHKCT  ) FIHKCT  ," '//変更件数
    sql = sql & " SUM(FIHKKG  ) FIHKKG  ," '//変更金額
    sql = sql & " SUM(FIKYCT  ) FIKYCT  ," '//解約件数
    sql = sql & " SUM(FIKYKG  ) FIKYKG  ," '//解約金額
    sql = sql & " SUM(ALLCNT_T) ALLCNT_T,"
    sql = sql & " SUM(FIHKCT_T) FIHKCT_T,"
    sql = sql & " SUM(FIHKKG_T) FIHKKG_T,"
    sql = sql & " SUM(FIKYCT_T) FIKYCT_T,"
    sql = sql & " SUM(FIKYKG_T) FIKYKG_T " & vbCrLf
    sql = sql & " FROM(" & vbCrLf
    '//保護者レコード
    sql = sql & " SELECT "
    sql = sql & " FIADID,"                                       '//処理結果
    sql = sql & " ABITCD,ABKJNM,FIKYCD,"
    sql = sql & " FIMKDT,"
    sql = sql & " 1                                     ALLCNT," '//件数
    sql = sql & " DECODE(NVL(FIKYFG,0),0,     1,     0) FIHKCT," '//変更件数    解約が０なら変更
    sql = sql & " DECODE(NVL(FIKYFG,0),0,FIHKKG,     0) FIHKKG," '//変更金額    解約が０なら変更
    sql = sql & " DECODE(NVL(FIKYFG,0),0,     0,     1) FIKYCT," '//解約件数
    sql = sql & " DECODE(NVL(FIKYFG,0),0,     0,FIHKKG) FIKYKG," '//解約金額
    sql = sql & " 0 ALLCNT_T,"
    sql = sql & " 0 FIHKCT_T,"
    sql = sql & " 0 FIHKKG_T,"
    sql = sql & " 0 FIKYCT_T,"
    sql = sql & " 0 FIKYKG_T "
    sql = sql & " FROM tfFurikaeYoteiImportTemp,"
    sql = sql & "      taItakushaMaster "
    sql = sql & "  WHERE FIITKB = ABITKB "
    sql = sql & "    AND FIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss')"
    sql = sql & "    AND FIRKBN <> -1 " & vbCrLf
    sql = sql & " UNION ALL " & vbCrLf      '// UNION ALL で無いといけない！！！
    '//合計レコード
    sql = sql & " SELECT "
    sql = sql & " FIADID,"                                              '//処理結果
    sql = sql & " ABITCD,ABKJNM,FIKYCD,"
    sql = sql & " FIMKDT FIMKDT,"
    sql = sql & " 0 ALLCNT,"
    sql = sql & " 0 FIHKCT,"
    sql = sql & " 0 FIHKKG,"
    sql = sql & " 0 FIKYCT,"
    sql = sql & " 0 FIKYKG,"
    sql = sql & " 1                                     ALLCNT_T," '//件数
    sql = sql & " NVL(FIHKCT,0)                         FIHKCT_T," '//変更件数
    sql = sql & " NVL(FIHKKG,0)                         FIHKKG_T," '//変更金額
    sql = sql & " NVL(FIKYCT,0)                         FIKYCT_T," '//解約件数
    sql = sql & " 0                                     FIKYKG_T " '//解約金額
    sql = sql & " FROM tfFurikaeYoteiImportTemp,"
    sql = sql & "      taItakushaMaster "
    sql = sql & "  WHERE FIITKB = ABITKB "
    sql = sql & "    AND FIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss')"
    sql = sql & "    AND FIRKBN = -1 " & vbCrLf
    sql = sql & ")" & vbCrLf
    sql = sql & " GROUP BY FIADID,ABITCD,ABKJNM,FIKYCD" & vbCrLf
    sql = sql & " ORDER BY FIADID,ABITCD,ABKJNM,FIKYCD"
    Dim reg As New RegistryClass
    Load rptYoteiReqImportReport
    With rptYoteiReqImportReport
        .lblCondition.Caption = lblImpDate.Caption & "：" & cboImpDate.Text
        .adoData.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=" & reg.DbPassword & _
                                    ";Persist Security Info=True;User ID=" & reg.DbUserName & _
                                                           ";Data Source=" & reg.DbDatabaseName
        .adoData.Source = sql
        'Call .adoData.Refresh
        Call .Show
    End With
    Set ms = Nothing
End Sub

Private Sub pImportDateRefresh()
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    sql = "SELECT DISTINCT FIINDT FROM tfFurikaeYoteiImportTemp"
    sql = sql & " WHERE FIINDT >= TO_DATE('" & txtStartDate.Text & "','yyyy/mm/dd hh24:mi:ss')"
    sql = sql & " ORDER BY FIINDT DESC "
    Set dyn = gdDBS.OpenRecordset(sql)
    cboImpDate.Clear
    Do Until dyn.EOF
        Call cboImpDate.AddItem(dyn.Fields("FIINDT").Value)
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    If 0 < cboImpDate.ListCount Then
        cboImpDate.ListIndex = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    Call mForm.LockedControl(False)
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    sql = "SELECT * FROM taSystemInformation"
    Set dyn = gdDBS.OpenRecordset(sql)
    If dyn.EOF Then
        mStartDate = Now()
    Else
        mStartDate = Format(dyn.Fields("AANWDT").Value, "yyyy/mm/dd hh:nn:ss")
    End If
    Call dyn.Close
    txtStartDate.Text = mStartDate
    chkDefault.Value = 1
    Call pImportDateRefresh
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmYoteiReqImportReport = Nothing
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

Private Sub txtStartDate_Change()
    Call pImportDateRefresh
End Sub
