VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmHolidayMaster 
   Caption         =   "休日マスタメンテナンス"
   ClientHeight    =   4665
   ClientLeft      =   2880
   ClientTop       =   2430
   ClientWidth     =   6555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6555
   Begin VB.ComboBox cboYear 
      Height          =   300
      Left            =   1440
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   18
      Top             =   900
      Width           =   735
   End
   Begin VB.Frame fraShoriKubun 
      Caption         =   "処理区分"
      Height          =   615
      Left            =   360
      TabIndex        =   13
      Top             =   120
      Width           =   2955
      Begin VB.OptionButton optShoriKubun 
         BackColor       =   &H000000FF&
         Caption         =   "参照"
         Height          =   255
         Index           =   3
         Left            =   2820
         TabIndex        =   19
         Tag             =   "InputKey"
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton optShoriKubun 
         Caption         =   "修正"
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optShoriKubun 
         Caption         =   "削除"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optShoriKubun 
         Caption         =   "新規"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblShoriKubun 
         BackColor       =   &H000000FF&
         Caption         =   "処理区分"
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   120
         Width           =   975
      End
   End
   Begin imText6Ctl.imText txtName 
      Height          =   285
      Left            =   2460
      TabIndex        =   2
      Top             =   3480
      Width           =   1695
      _Version        =   65537
      _ExtentX        =   2990
      _ExtentY        =   503
      Caption         =   "休日マスタメンテナンス.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "休日マスタメンテナンス.frx":006E
      Key             =   "休日マスタメンテナンス.frx":008C
      MouseIcon       =   "休日マスタメンテナンス.frx":00D0
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
      MaxLength       =   14
      LengthAsByte    =   -1
      Text            =   "休日名称"
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   4
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin imDate6Ctl.imDate txtHoliday 
      Height          =   285
      Left            =   1380
      TabIndex        =   1
      Top             =   3480
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   503
      Calendar        =   "休日マスタメンテナンス.frx":00EC
      Caption         =   "休日マスタメンテナンス.frx":0272
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "休日マスタメンテナンス.frx":02E0
      Keys            =   "休日マスタメンテナンス.frx":02FE
      MouseIcon       =   "休日マスタメンテナンス.frx":035C
      Spin            =   "休日マスタメンテナンス.frx":0378
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
      HighlightText   =   2
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
      Text            =   "2002/11/18"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   37578
      CenturyMode     =   0
   End
   Begin ORADCLibCtl.ORADC dbcHoliday 
      Height          =   315
      Left            =   2700
      Top             =   4020
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   207
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   "dcssvr03"
      Connect         =   "kumon/kumon"
      RecordSource    =   $"休日マスタメンテナンス.frx":03A0
      ReadOnly        =   -1  'True
   End
   Begin MSDBCtls.DBList dblHoliday 
      Bindings        =   "休日マスタメンテナンス.frx":045F
      DataField       =   "HOLIDAY"
      DataSource      =   "dbcHoliday"
      Height          =   1800
      Left            =   1380
      TabIndex        =   0
      Top             =   1620
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3175
      _Version        =   393216
      IntegralHeight  =   0   'False
      ListField       =   "HOLIDAY"
      BoundColumn     =   "HOLIDAY"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cboKubun 
      Height          =   300
      ItemData        =   "休日マスタメンテナンス.frx":0484
      Left            =   4200
      List            =   "休日マスタメンテナンス.frx":0491
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   3
      Top             =   3480
      Width           =   1635
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "更新(&U)"
      Height          =   435
      Left            =   660
      TabIndex        =   4
      Top             =   3960
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "終了(&X)"
      Height          =   435
      Left            =   4740
      TabIndex        =   5
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   60
      Width           =   1395
   End
   Begin VB.Label Label5 
      Caption         =   "年"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   960
      Width           =   315
   End
   Begin VB.Label Label4 
      Caption         =   "対象年度"
      Height          =   255
      Left            =   540
      TabIndex        =   10
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      Caption         =   "種別"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4500
      TabIndex        =   9
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      Caption         =   "休日名称"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1380
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      Caption         =   "年月日"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1380
      TabIndex        =   7
      Top             =   1380
      Width           =   1155
   End
   Begin VB.Label Label8 
      Caption         =   "祝日設定日"
      Height          =   255
      Left            =   420
      TabIndex        =   6
      Top             =   3540
      Width           =   915
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
Attribute VB_Name = "frmHolidayMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm As New FormClass
Private mCaption As String

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    Dim sql As String, cnt As Long
    Select Case lblShoriKubun.Caption
    Case eShoriKubun.Add, eShoriKubun.Edit
        sql = "UPDATE teHolidayMaster SET "
        sql = sql & " EANAME = '" & txtName.Text & "',"
        sql = sql & " EAHDKB = " & cboKubun.ItemData(cboKubun.ListIndex)
        sql = sql & " WHERE EADATE = " & txtHoliday.Number
        cnt = gdDBS.Database.ExecuteSQL(sql)
        '//追加時でかつレコード無しのみの動作
        If lblShoriKubun.Caption = eShoriKubun.Add And cnt = 0 Then
            sql = "INSERT INTO teHolidayMaster "
            sql = sql & " VALUES("
            sql = sql & txtHoliday.Number & ","
            sql = sql & "'" & txtName.Text & "',"
            sql = sql & cboKubun.ItemData(cboKubun.ListIndex)
            sql = sql & ")"
            Call gdDBS.Database.ExecuteSQL(sql)
        End If
    Case eShoriKubun.Delete
        sql = "DELETE teHolidayMaster "
        sql = sql & " WHERE EADATE = " & txtHoliday.Number
        Call gdDBS.Database.ExecuteSQL(sql)
    End Select
    Call cboYear_Click  '//変更内容を反映するために
End Sub

Private Sub dblHoliday_Click()
    txtHoliday.Text = Left(dblHoliday.Text, Len("2002/07/10"))
    If "" <> Trim(txtHoliday.Text) Then
#If ORA_DEBUG = 1 Then
        Dim sql As String, dyn As OraDynaset
#Else
        Dim sql As String, dyn As Object
#End If
        sql = "SELECT * FROM teHolidayMaster"
        sql = sql & " WHERE eadate = " & txtHoliday.Number
        Set dyn = gdDBS.OpenRecordset(sql)
        txtName.Text = Trim(gdDBS.Nz(dyn.Fields("eaname")))
        cboKubun.ListIndex = Val(gdDBS.Nz(dyn.Fields("eahdkb")))
    End If
End Sub

Private Sub pCheckAndInsert(vYMD As Long, vName As Variant, vHoliday As Integer)
#If ORA_DEBUG = 1 Then
        Dim sql As String, dyn As OraDynaset
#Else
        Dim sql As String, dyn As Object
#End If
    sql = "SELECT * FROM teHolidayMaster"
    sql = sql & " WHERE EADATE = TO_CHAR(TO_DATE(" & vYMD & ",'YYYYMMDD') + " & Abs(vHoliday <> 0) & ",'YYYYMMDD')"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If Not dyn.EOF() Then
        Exit Sub
    End If
    sql = "INSERT INTO teHolidayMaster VALUES("
    sql = sql & "TO_CHAR(TO_DATE(" & vYMD & ",'YYYYMMDD') + " & Abs(vHoliday <> 0) & ",'YYYYMMDD'),"
    sql = sql & "'" & vName & "',"
    sql = sql & "'" & vHoliday & "'"
    sql = sql & ")"
    Call gdDBS.Database.ExecuteSQL(sql)
End Sub

Private Sub pMakeHoliday(vYear As Integer)
#If ORA_DEBUG = 1 Then
        Dim sql As String, dyn As OraDynaset
#Else
        Dim sql As String, dyn As Object
#End If
    sql = "SELECT * FROM teHolidayMaster"
    sql = sql & " WHERE eadate BETWEEN " & vYear & "0101 AND " & vYear & "1231"
    Set dyn = gdDBS.OpenRecordset(sql)
'    Set dyn = dbcHoliday.Database.CreateDynaset(sql, 0)
    If dyn.RecordCount <> 0 Then
        Exit Sub
    End If
'//2002/10/10 現在で固定の祝日を定義
    Dim DateTable As Variant, NameTable As Variant, i As Integer
    DateTable = Array("0101", "0211", "0429", "0503", "0504", "0505", "0720", "0915", "1103", "1123", "1223")
    NameTable = Array("元旦", "建国記念日", "みどりの日", "憲法記念日", "国民の休日", "子供の日", "海の日", "敬老の日", "文化の日", "勤労感謝の日", "天皇誕生日")
    Call gdDBS.Database.BeginTrans
    '//最後からしないと 5/3,4,5 の振替休日が変になる
    For i = UBound(DateTable) To LBound(DateTable) Step -1
        Call pCheckAndInsert(vYear & DateTable(i), NameTable(i), 0)     '0=祝日
        '//祝日が日曜日の時は振替休日を作成
        If Weekday(DateSerial(vYear, Left(DateTable(i), 2), Right(DateTable(i), 2))) = vbSunday Then
            Call pCheckAndInsert(vYear & DateTable(i), NameTable(i), 1) '1=振替休日
        End If
    Next i
    Call gdDBS.Database.CommitTrans
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    Call mForm.LockedControl(False)
    optShoriKubun(eShoriKubun.Edit).Value = True    '//ここで参照はしない
    dbcHoliday.ReadOnly = True  'リスト内のデータは更新しない
    '//祝日区分コンボ設定
    Call cboKubun.Clear
    Call cboKubun.AddItem("祝日"):       cboKubun.ItemData(0) = 0
    Call cboKubun.AddItem("振替休日"):   cboKubun.ItemData(1) = 1
    Call cboKubun.AddItem("その他"):     cboKubun.ItemData(2) = 2
    
    '//処理できる年度をコンボボックスに設定する 当年基準
    Dim i As Integer
    Call cboYear.Clear
    For i = Year(Now()) - 1 To Year(Now()) + 1
        Call cboYear.AddItem(i)
        If i = Year(Now()) Then
            cboYear.ListIndex = cboYear.NewIndex
        End If
    Next i
    Call cboYear_Click
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmHolidayMaster = Nothing
    Set mForm = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub optShoriKubun_Click(Index As Integer)
    lblShoriKubun.Caption = Index
    Select Case Index
    Case eShoriKubun.Add
        txtHoliday.Enabled = True
        txtName.Enabled = True
        cboKubun.Enabled = True
    Case eShoriKubun.Edit
        txtHoliday.Enabled = False
        txtName.Enabled = True
        cboKubun.Enabled = True
    Case eShoriKubun.Delete
        txtHoliday.Enabled = False
        txtName.Enabled = False
        cboKubun.Enabled = False
    End Select
End Sub

Private Sub cboYear_Click()
    Dim sql As String
    sql = "SELECT TO_CHAR(TO_DATE(eadate,'YYYYMMDD'),'YYYY/MM/DD') || ' ' "
    sql = sql & " || SUBSTRB(eaname || '　　　　　　　　　　　　　　',1,20)"
    sql = sql & " || DECODE(eahdkb,'0','祝日','1','振替休日','その他') AS Holiday"
    sql = sql & " FROM teHolidayMaster"
    sql = sql & " WHERE eadate BETWEEN " & cboYear.Text & "0101 AND " & cboYear.Text & "1231"
    sql = sql & " ORDER BY EADATE"
    dbcHoliday.RecordSource = sql
    dbcHoliday.Refresh
    dblHoliday.ListField = "Holiday"
    If 0 = dbcHoliday.Recordset.RecordCount Then
        Call pMakeHoliday(cboYear.Text)
        Call cboYear_Click
    End If
#If ORA_DEBUG = 1 Then
    Dim dyn As OraDynaset
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Dim dyn As Object
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If Not dyn.EOF() Then
        dblHoliday.BoundText = dyn.Fields("Holiday")
    End If
    Set dyn = Nothing
End Sub

Private Sub mnuEnd_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

Private Sub txtHoliday_DropOpen(NoDefault As Boolean)
    txtHoliday.Calendar.Holidays = gdDBS.Holiday(txtHoliday.Year)
End Sub

