VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmKeiyakushaNayose 
   Caption         =   "名寄せオーナー 一覧"
   ClientHeight    =   5520
   ClientLeft      =   8640
   ClientTop       =   2160
   ClientWidth     =   9705
   KeyPreview      =   -1  'True
   LinkTopic       =   "名寄せ"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9705
   Begin VB.CommandButton cmdSelect 
      Caption         =   "選択(&S)"
      Height          =   495
      Left            =   300
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "閉じる(&C)"
      Height          =   495
      Left            =   8025
      TabIndex        =   3
      Top             =   4800
      Width           =   1335
   End
   Begin MSDBCtls.DBList dblNayoseList 
      Bindings        =   "契約者名寄せ一覧.frx":0000
      DataField       =   "NAYOSE_LIST"
      Height          =   3300
      Left            =   300
      TabIndex        =   0
      Top             =   975
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   5821
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin imText6Ctl.imText txtBAKYCD 
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   300
      TabIndex        =   1
      Tag             =   "InputKey"
      Top             =   375
      Width           =   795
      _Version        =   65537
      _ExtentX        =   1402
      _ExtentY        =   503
      Caption         =   "契約者名寄せ一覧.frx":0022
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "契約者名寄せ一覧.frx":008E
      Key             =   "契約者名寄せ一覧.frx":00AC
      MouseIcon       =   "契約者名寄せ一覧.frx":00F0
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
      AllowSpace      =   0
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
   Begin ORADCLibCtl.ORADC dbcKeiyakushaMaster 
      Height          =   315
      Left            =   1800
      Top             =   4800
      Visible         =   0   'False
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
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
      DatabaseName    =   "dbsvr03"
      Connect         =   "wao/wao"
      RecordSource    =   "SELECT BAKYCD || ':' || BAKJNM as NAYOSE_LIST  FROM tbKeiyakushaMaster"
   End
   Begin VB.Label Label6 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "　オーナー名(漢字)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5600
      TabIndex        =   11
      Top             =   750
      Width           =   3765
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "(名寄せ先)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1025
      TabIndex        =   10
      Top             =   750
      Width           =   950
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "オーナー"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   300
      TabIndex        =   8
      Top             =   750
      Width           =   750
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "「選択(&S)」ボタンで画面の選択されたオーナーを修正可能です。"
      Height          =   180
      Left            =   300
      TabIndex        =   7
      Tag             =   "InputKey"
      Top             =   4425
      Width           =   4785
   End
   Begin VB.Label lblBAKJNM 
      Caption         =   "契約者名"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   405
      Width           =   3720
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   255
      Left            =   3900
      TabIndex        =   5
      Top             =   0
      Width           =   1395
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "名寄せオーナー番号"
      Height          =   180
      Left            =   300
      TabIndex        =   4
      Top             =   75
      Width           =   1590
   End
   Begin VB.Label Label4 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "　校名(漢字)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1950
      TabIndex        =   9
      Top             =   750
      Width           =   3660
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
Attribute VB_Name = "frmKeiyakushaNayose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mForm As New FormClass
Private mCaption As String
Public m_BAITKB As String
Public m_Params As String
Public m_Result As String

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    
    '//受取りパラメータの処理
    Dim code() As String
    code = Split(m_Params, vbTab)
    m_BAITKB = code(0)
    Call makeNayoseList(m_BAITKB, code(1))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Unload(cancel As Integer)
    Set frmKeiyakushaNayose = Nothing
    Set mForm = Nothing
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        cancel = True
    End If
End Sub

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    On Error Resume Next
    Dim result() As String
    result = Split(dblNayoseList.Text, vbTab)
    m_Result = result(0)
    Unload Me
End Sub

Private Sub dblNayoseList_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub makeNayoseList(vBAITKB As String, vBAKYCD As String)
    Dim sql As String, dyn As Object
    
    '//オーナー名の表示
    sql = "select * from tbKeiyakushaMaster"
    sql = sql & " where BAITKB = '" & vBAITKB & "'"
    sql = sql & "   and BAKYCD = '" & vBAKYCD & "'"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    Dim name As String
    If False = dyn.EOF Then
        name = dyn.Fields("BAKJNM")
    End If
    txtBAKYCD.Text = vBAKYCD
    lblBAKJNM.Caption = name
    
    '//名寄せ名の表示
    sql = "SELECT "
    sql = sql & " BAKYCD ||chr(9)|| ' ('"
    sql = sql & "|| NVL(BAKYNY,BAKYCD) "
    sql = sql & "|| ') '"
    sql = sql & "|| SUBSTR(RPAD(BAKOME,40,' '),1,40) "
    sql = sql & "|| SUBSTR(RPAD(BAKJNM,40,' '),1,40) "
    sql = sql & " as NAYOSE_LIST "
    sql = sql & " FROM tbKeiyakushaMaster m "
    sql = sql & " where BAITKB = '" & vBAITKB & "'"
#If 1 Then
    sql = sql & "   and BAKYNY is not null"
#End If
    '//オーナー番号がある場合のみ
    If vBAKYCD <> "" Then
        sql = sql & "   and BAKYNY = '" & vBAKYCD & "'"
    End If
    sql = sql & "   and (BAITKB,BAKYCD,BASQNO) in("
    sql = sql & "       select BAITKB,BAKYCD,MAX(BASQNO) from tbKeiyakushaMaster s "
    sql = sql & "       where s.BAITKB = m.BAITKB "
    sql = sql & "         and s.BAKYCD = m.BAKYCD "
    sql = sql & "       group by BAITKB,BAKYCD"
    sql = sql & "   )"
    sql = sql & " ORDER BY LTRIM(nvl(BAKYNY,'XXX')),BAKYCD"
    dbcKeiyakushaMaster.RecordSource = sql
    'dbcKeiyakushaMaster.ReadOnly = True
    dbcKeiyakushaMaster.Refresh
    dblNayoseList.ListField = "NAYOSE_LIST"
End Sub

Private Sub txtBAKYCD_LostFocus()
    If Trim(txtBAKYCD.Text) <> "" Then
        txtBAKYCD.Text = Format(Val(txtBAKYCD.Text), String(7, "0"))
    End If
    Call makeNayoseList(m_BAITKB, txtBAKYCD.Text)
End Sub

Private Sub mnuEnd_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub
