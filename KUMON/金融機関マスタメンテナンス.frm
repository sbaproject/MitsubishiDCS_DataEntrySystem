VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmBankMaster 
   Caption         =   "���Z�@�փ}�X�^�����e�i���X"
   ClientHeight    =   7410
   ClientLeft      =   3855
   ClientTop       =   3855
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   9120
   Begin VB.Frame fraOldNew 
      Caption         =   "�V�E���Z�@��"
      Height          =   1515
      Index           =   1
      Left            =   4140
      TabIndex        =   7
      Tag             =   "InputKey"
      Top             =   840
      Width           =   3255
      Begin imText6Ctl.imText txtDASITN 
         DataField       =   "CASITN"
         Height          =   285
         Index           =   1
         Left            =   1140
         TabIndex        =   9
         Tag             =   "InputKey"
         Top             =   660
         Width           =   375
         _Version        =   65537
         _ExtentX        =   661
         _ExtentY        =   503
         Caption         =   "���Z�@�փ}�X�^�����e�i���X.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "���Z�@�փ}�X�^�����e�i���X.frx":006E
         Key             =   "���Z�@�փ}�X�^�����e�i���X.frx":008C
         MouseIcon       =   "���Z�@�փ}�X�^�����e�i���X.frx":00D0
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
         MaxLength       =   3
         LengthAsByte    =   -1
         Text            =   "123"
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
      Begin imText6Ctl.imText txtDABANK 
         DataField       =   "CABANK"
         Height          =   285
         Index           =   1
         Left            =   1140
         TabIndex        =   8
         Tag             =   "InputKey"
         Top             =   300
         Width           =   495
         _Version        =   65537
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   "���Z�@�փ}�X�^�����e�i���X.frx":00EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "���Z�@�փ}�X�^�����e�i���X.frx":015A
         Key             =   "���Z�@�փ}�X�^�����e�i���X.frx":0178
         MouseIcon       =   "���Z�@�փ}�X�^�����e�i���X.frx":01BC
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
         MaxLength       =   4
         LengthAsByte    =   -1
         Text            =   "1234"
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
      Begin imDate6Ctl.imDate txtDAYKED 
         DataField       =   "CAKYST"
         Height          =   315
         Index           =   1
         Left            =   1140
         TabIndex        =   10
         Tag             =   "InputKey"
         Top             =   1020
         Width           =   1275
         _Version        =   65537
         _ExtentX        =   2249
         _ExtentY        =   556
         Calendar        =   "���Z�@�փ}�X�^�����e�i���X.frx":01D8
         Caption         =   "���Z�@�փ}�X�^�����e�i���X.frx":035E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "���Z�@�փ}�X�^�����e�i���X.frx":03CC
         Keys            =   "���Z�@�փ}�X�^�����e�i���X.frx":03EA
         MouseIcon       =   "���Z�@�փ}�X�^�����e�i���X.frx":0448
         Spin            =   "���Z�@�փ}�X�^�����e�i���X.frx":0464
         AlignHorizontal =   2
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   1
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
         Text            =   "    /  /  "
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   -2
         CenturyMode     =   0
      End
      Begin VB.Label lblShitenName 
         Caption         =   "�x�X��"
         Height          =   255
         Index           =   1
         Left            =   1740
         TabIndex        =   32
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblBankName 
         Caption         =   "��s��"
         Height          =   255
         Index           =   1
         Left            =   1740
         TabIndex        =   30
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label lblBankcode 
         Alignment       =   1  '�E����
         Caption         =   "���Z�@��"
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   28
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblShitenCode 
         Alignment       =   1  '�E����
         Caption         =   "�x�X"
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   27
         Tag             =   "InputKey"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblTekiyoBi 
         Alignment       =   1  '�E����
         Caption         =   "�K�p�J�n��"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Tag             =   "InputKey"
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label lblBikou 
         Caption         =   "���"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   25
         Tag             =   "InputKey"
         Top             =   1080
         Width           =   435
      End
   End
   Begin VB.Frame fraOldNew 
      Caption         =   "���E���Z�@��"
      Height          =   1515
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   3255
      Begin imText6Ctl.imText txtDASITN 
         DataField       =   "CASITN"
         Height          =   285
         Index           =   0
         Left            =   1140
         TabIndex        =   5
         Tag             =   "InputKey"
         Top             =   660
         Width           =   375
         _Version        =   65537
         _ExtentX        =   661
         _ExtentY        =   503
         Caption         =   "���Z�@�փ}�X�^�����e�i���X.frx":048C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "���Z�@�փ}�X�^�����e�i���X.frx":04FA
         Key             =   "���Z�@�փ}�X�^�����e�i���X.frx":0518
         MouseIcon       =   "���Z�@�փ}�X�^�����e�i���X.frx":055C
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
         MaxLength       =   3
         LengthAsByte    =   -1
         Text            =   "123"
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
      Begin imText6Ctl.imText txtDABANK 
         DataField       =   "CABANK"
         Height          =   285
         Index           =   0
         Left            =   1140
         TabIndex        =   4
         Tag             =   "InputKey"
         Top             =   300
         Width           =   495
         _Version        =   65537
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   "���Z�@�փ}�X�^�����e�i���X.frx":0578
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "���Z�@�փ}�X�^�����e�i���X.frx":05E6
         Key             =   "���Z�@�փ}�X�^�����e�i���X.frx":0604
         MouseIcon       =   "���Z�@�փ}�X�^�����e�i���X.frx":0648
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
         MaxLength       =   4
         LengthAsByte    =   -1
         Text            =   "1234"
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
      Begin imDate6Ctl.imDate txtDAYKED 
         DataField       =   "CAKYST"
         Height          =   315
         Index           =   0
         Left            =   1140
         TabIndex        =   6
         Tag             =   "InputKey"
         Top             =   1020
         Width           =   1275
         _Version        =   65537
         _ExtentX        =   2249
         _ExtentY        =   556
         Calendar        =   "���Z�@�փ}�X�^�����e�i���X.frx":0664
         Caption         =   "���Z�@�փ}�X�^�����e�i���X.frx":07EA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "���Z�@�փ}�X�^�����e�i���X.frx":0858
         Keys            =   "���Z�@�փ}�X�^�����e�i���X.frx":0876
         MouseIcon       =   "���Z�@�փ}�X�^�����e�i���X.frx":08D4
         Spin            =   "���Z�@�փ}�X�^�����e�i���X.frx":08F0
         AlignHorizontal =   2
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   1
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
         Text            =   "    /  /  "
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   -2
         CenturyMode     =   0
      End
      Begin VB.Label lblShitenName 
         Caption         =   "�x�X��"
         Height          =   255
         Index           =   0
         Left            =   1740
         TabIndex        =   31
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblBankName 
         Caption         =   "��s��"
         Height          =   255
         Index           =   0
         Left            =   1740
         TabIndex        =   29
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label lblBikou 
         Caption         =   "�܂�"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   24
         Tag             =   "InputKey"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblBankcode 
         Alignment       =   1  '�E����
         Caption         =   "���Z�@��"
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblShitenCode 
         Alignment       =   1  '�E����
         Caption         =   "�x�X"
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   22
         Tag             =   "InputKey"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblTekiyoBi 
         Alignment       =   1  '�E����
         Caption         =   "�K�p�I����"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Tag             =   "InputKey"
         Top             =   1080
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "�ΏێҌ���(&S)"
      Height          =   435
      Left            =   420
      TabIndex        =   11
      Top             =   2460
      Width           =   1455
   End
   Begin VB.Frame fraUpdateKubun 
      Caption         =   "�����敪"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2835
      Begin VB.OptionButton optShoriKubun 
         Caption         =   "�p�~"
         Height          =   255
         Index           =   1
         Left            =   1860
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optShoriKubun 
         Caption         =   "�����E���p��"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label lblShoriKubun 
         BackColor       =   &H000000FF&
         Caption         =   "�����敪"
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�X�V(&U)"
      Height          =   435
      Left            =   1020
      TabIndex        =   13
      Top             =   6720
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "�I��(&X)"
      Height          =   435
      Left            =   7140
      TabIndex        =   14
      Top             =   6720
      Width           =   1335
   End
   Begin ORADCLibCtl.ORADC dbcTrans 
      Height          =   315
      Left            =   4200
      Top             =   6660
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   207
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   "dcssvr03"
      Connect         =   "kumon/kumon"
      RecordSource    =   $"���Z�@�փ}�X�^�����e�i���X.frx":0918
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "���Z�@�փ}�X�^�����e�i���X.frx":09BF
      Height          =   3410
      Left            =   420
      OleObjectBlob   =   "���Z�@�փ}�X�^�����e�i���X.frx":09D6
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3000
      Width           =   8325
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   255
      Left            =   7680
      TabIndex        =   19
      Top             =   60
      Width           =   1395
   End
   Begin VB.Label lblKouzaCount 
      Alignment       =   1  '�E����
      Caption         =   "5,678"
      Height          =   195
      Left            =   2880
      TabIndex        =   18
      Top             =   2700
      Width           =   915
   End
   Begin VB.Label Label11 
      Alignment       =   1  '�E����
      Caption         =   "�g�p������"
      Height          =   195
      Left            =   2880
      TabIndex        =   17
      Top             =   2460
      Width           =   915
   End
   Begin VB.Label lblBankCount 
      Alignment       =   1  '�E����
      Caption         =   "1,234"
      Height          =   195
      Left            =   1980
      TabIndex        =   16
      Top             =   2700
      Width           =   795
   End
   Begin VB.Label Label7 
      Alignment       =   1  '�E����
      Caption         =   "�Y������"
      Height          =   195
      Left            =   1980
      TabIndex        =   15
      Top             =   2460
      Width           =   795
   End
   Begin VB.Menu mnuFile 
      Caption         =   "̧��(&F)"
      Begin VB.Menu mnuEnd 
         Caption         =   "�I��(&X)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuVersion 
         Caption         =   "�ް�ޮݏ��(&A)"
      End
   End
End
Attribute VB_Name = "frmBankMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm As New FormClass
Private mCaption As String

Private Enum eOldNew
    OldBank = 0
    NewBank = 1
End Enum

Private Enum eShoriKubun
    TouHaigou = 0
    Haishi = 1
End Enum

Private Enum eTable
    Itakusha = 0
    Hogosha = 1
End Enum

Private Sub pLockedControl(blMode As Boolean)
    'Call mForm.LockedControl(blMode)
    txtDABANK(eOldNew.OldBank).Text = ""
    txtDABANK(eOldNew.NewBank).Text = ""
    txtDASITN(eOldNew.OldBank).Text = ""
    txtDASITN(eOldNew.NewBank).Text = ""
    txtDAYKED(eOldNew.OldBank).Number = 0
    txtDAYKED(eOldNew.NewBank).Number = 0
    txtDAYKED(eOldNew.NewBank).Enabled = False
    lblBankCount.Caption = 0
    lblKouzaCount.Caption = 0
    lblBankName(eOldNew.OldBank).Caption = ""
    lblBankName(eOldNew.NewBank).Caption = ""
    lblShitenName(eOldNew.OldBank).Caption = ""
    lblShitenName(eOldNew.NewBank).Caption = ""
    cmdEnd.Enabled = True
    cmdUpdate.Enabled = True
End Sub

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Function pCheckBank(vBank As String, Optional vShiten As String = "") As Boolean
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    sql = "SELECT DABANK FROM tdBankMaster"
    sql = sql & " WHERE DABANK = '" & vBank & "'"
    If "" <> vShiten Then
        sql = sql & " AND DASITN = '" & vShiten & "'"
    End If
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
    pCheckBank = Not dyn.EOF
End Function

Private Function pInputCheck(Optional ByVal vMode As Boolean = False) As Boolean
    Dim obj As Object, msg As String
    If "" = txtDABANK(eOldNew.OldBank).Text Then
        msg = fraOldNew(eOldNew.OldBank).Caption & "�́u" & lblBankcode(eOldNew.OldBank).Caption & "�v�͕K�{���͂ł�."
        Set obj = txtDABANK(eOldNew.OldBank)
    ElseIf False = pCheckBank(txtDABANK(eOldNew.OldBank).Text, txtDASITN(eOldNew.OldBank).Text) Then
        msg = fraOldNew(eOldNew.OldBank).Caption & "�́u" & lblBankcode(eOldNew.OldBank).Caption & "�v�͑��݂��܂���."
        Set obj = txtDABANK(eOldNew.OldBank)
    ElseIf IsNull(txtDAYKED(eOldNew.OldBank).Number) Or txtDAYKED(eOldNew.OldBank).Number = 0 Then
        msg = fraOldNew(eOldNew.OldBank).Caption & "�́u" & lblTekiyoBi(eOldNew.OldBank).Caption & "�v�͕K�{���͂ł�."
        Set obj = txtDAYKED(eOldNew.OldBank)
    '//�X�V�� vMode = True
    ElseIf vMode = True Then
        If txtDABANK(eOldNew.OldBank).Text = txtDABANK(eOldNew.NewBank).Text _
        And "" = txtDASITN(eOldNew.OldBank).Text And "" = txtDASITN(eOldNew.NewBank).Text Then
            msg = "�V�E���ł̓���" & lblBankcode(eOldNew.OldBank).Caption & "�͐ݒ�ł��܂���."
            Set obj = txtDABANK(eOldNew.OldBank)
        ElseIf txtDABANK(eOldNew.OldBank).Text = txtDABANK(eOldNew.NewBank).Text _
           And txtDASITN(eOldNew.OldBank).Text = txtDASITN(eOldNew.NewBank).Text Then
            msg = "�V�E���ł̓���" & lblShitenCode(eOldNew.OldBank).Caption & "�͐ݒ�ł��܂���."
            Set obj = txtDASITN(eOldNew.OldBank)
        End If
        Select Case lblShoriKubun.Caption
        Case eShoriKubun.TouHaigou
            If "" = txtDABANK(eOldNew.NewBank).Text Then
                msg = fraOldNew(eOldNew.NewBank).Caption & "�́u" & lblBankcode(eOldNew.NewBank).Caption & "�v�͕K�{���͂ł�."
                Set obj = txtDABANK(eOldNew.NewBank)
            ElseIf False = pCheckBank(txtDABANK(eOldNew.NewBank).Text, txtDASITN(eOldNew.NewBank).Text) Then
                msg = fraOldNew(eOldNew.NewBank).Caption & "�́u" & lblBankcode(eOldNew.NewBank).Caption & "�v�͑��݂��܂���."
                Set obj = txtDABANK(eOldNew.NewBank)
            ElseIf "" <> txtDASITN(eOldNew.OldBank).Text And "" = txtDASITN(eOldNew.NewBank).Text Then
                msg = fraOldNew(eOldNew.NewBank).Caption & "�́u" & lblShitenCode(eOldNew.NewBank).Caption & "�v�͕K�{���͂ł�."
                Set obj = txtDASITN(eOldNew.NewBank)
            End If
        Case eShoriKubun.Haishi
        End Select
    End If
    If TypeName(obj) <> "Nothing" Then
        Call MsgBox(msg, vbOKOnly, mCaption)
        Call obj.SetFocus
        Exit Function
    End If
    pInputCheck = True
End Function

Private Sub cmdSearch_Click()
    If False = pInputCheck() Then
        Exit Sub
    End If
    Dim sql As String
    Dim ms As New MouseClass
    Call ms.Start
    sql = ""
    sql = sql & "SELECT 0 OrderKey," & vbCrLf
    sql = sql & "'�_���' Kubun," & vbCrLf
    sql = sql & "BAKYCD Code1," & vbCrLf
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//    sql = sql & "BAKSCD Code2," & vbCrLf
    sql = sql & "NULL   Code2," & vbCrLf
    sql = sql & "NULL   Code3," & vbCrLf
    sql = sql & "BASQNO SeqNo," & vbCrLf
    sql = sql & "BAKJNM Name," & vbCrLf
    sql = sql & "BABANK Bank," & vbCrLf
    sql = sql & "BASITN Shiten," & vbCrLf
    sql = sql & "DECODE(BAKZSB,'1','����','2','����') Shubetsu," & vbCrLf
    sql = sql & "BAKZNO KouzaNo" & vbCrLf
    sql = sql & " FROM tbKeiyakushaMaster" & vbCrLf
    sql = sql & pMakeWhereSQL(eTable.Itakusha)
    sql = sql & " UNION ALL " & vbCrLf
    sql = sql & "SELECT 1 OrderKey," & vbCrLf
    sql = sql & "'�ی��' Kubun," & vbCrLf
    sql = sql & "CAKYCD Code1," & vbCrLf
    sql = sql & "CAKSCD Code2," & vbCrLf
    sql = sql & "CAHGCD Code3," & vbCrLf
    sql = sql & "CASQNO SeqNo," & vbCrLf
    sql = sql & "CAKJNM Name," & vbCrLf
    sql = sql & "CABANK Bank," & vbCrLf
    sql = sql & "CASITN Shiten," & vbCrLf
    sql = sql & "DECODE(CAKZSB,'1','����','2','����') Shubetsu," & vbCrLf
    sql = sql & "CAKZNO KouzaNo" & vbCrLf
    sql = sql & " FROM tcHogoshaMaster" & vbCrLf
    sql = sql & pMakeWhereSQL(eTable.Hogosha)
    sql = sql & " ORDER BY OrderKey,Code1,Code2,Code3,SeqNo"
    dbcTrans.RecordSource = sql
    dbcTrans.Refresh
    lblKouzaCount.Caption = dbcTrans.Recordset.RecordCount
    If dbcTrans.Recordset.RecordCount = 0 Then
        Call MsgBox("�Ώێ҃f�[�^�͑��݂��܂���.", vbInformation, mCaption)
        Exit Sub
    End If
End Sub

Private Sub cmdUpdate_Click()
    If False = pInputCheck(True) Then
        Exit Sub
    End If
    If vbOK <> MsgBox("�Y������u�_��ҁv�Ɓu�ی�ҁv�̋��Z�@�֏���ǉ����܂��B" & vbCrLf & "��낵���ł����H", vbInformation + vbOKCancel + vbDefaultButton2, mCaption) Then
        Exit Sub
    End If
    Dim ms As New MouseClass
    Call ms.Start
    
    Call pMakeNewRecord
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    dbcTrans.RecordSource = ""
    dbcTrans.ReadOnly = True
    Call pLockedControl(True)
    optShoriKubun(eOldNew.OldBank).Value = True
    lblShoriKubun.Caption = 0
    txtDAYKED(eOldNew.OldBank).MinDate = gdDBS.sysDate("YYYY/MM/DD")
    txtDAYKED(eOldNew.NewBank).MinDate = gdDBS.sysDate("YYYY/MM/DD")
    '//MinDate ��ݒ肷��� .Number �̒l�����̒l�ɐݒ肳��Ă��܂��̂ōď�����
    txtDAYKED(eOldNew.OldBank).Number = 0
    txtDAYKED(eOldNew.NewBank).Number = 0
'    Call txtDABANK(eoldnew.oldbank).SetFocus
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBankMaster = Nothing
    Set mForm = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub optShoriKubun_Click(Index As Integer)
    Dim flag As Boolean
    flag = Index <> eShoriKubun.Haishi
    lblShoriKubun.Caption = Index
    fraOldNew(eOldNew.NewBank).Enabled = flag
    lblBankcode(eOldNew.NewBank).Enabled = flag
    lblBankName(eOldNew.NewBank).Enabled = flag
    txtDABANK(eOldNew.NewBank).Enabled = flag
    lblShitenCode(eOldNew.NewBank).Enabled = flag
    lblShitenName(eOldNew.NewBank).Enabled = flag
    txtDASITN(eOldNew.NewBank).Enabled = flag
    lblTekiyoBi(eOldNew.NewBank).Enabled = flag
    lblBikou(eOldNew.NewBank).Enabled = flag
    txtDAYKED(eOldNew.NewBank).Enabled = False    '//����͏�� ��\��
    '//�V�E���Z�@�֏��͏�ɏ�����
    txtDABANK(eOldNew.NewBank).Text = ""
    txtDASITN(eOldNew.NewBank).Text = ""
    txtDAYKED(eOldNew.NewBank).Number = 0
    lblBankName(eOldNew.NewBank).Caption = ""
    lblShitenName(eOldNew.NewBank).Caption = ""
End Sub

#If 0 Then
Private Sub txtDABANK_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not (KeyCode = vbKeyReturn) Then
        Exit Sub
    End If
    
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim msg As String
        
    lblBankName(Index).Caption = ""
    txtDASITN(Index).Text = ""
    lblShitenName(Index).Caption = ""
    If "" = Trim(txtDABANK(Index).Text) Then
        Exit Sub
    End If
'''2002/10/09 �z�X�g�f�[�^�̊֌W�Ńt�B�[���h���폜����
'''    sql = "SELECT DAKJNM,DAYKED FROM tdBankMaster" & vbCrLf
    sql = "SELECT DAKJNM FROM tdBankMaster" & vbCrLf
    sql = sql & " WHERE DARKBN = '" & eBankRecordKubun.Bank & "'" & vbCrLf
    sql = sql & "   AND DABANK = '" & txtDABANK(Index).Text & "'" & vbCrLf
'''2002/10/09 �z�X�g�f�[�^�̊֌W�Ńt�B�[���h���폜����
'''    sql = sql & "   AND TO_CHAR(SYSDATE,'YYYYMMDD') BETWEEN DAYKST AND DAYKED" & vbCrLf
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    '//���E���Z�@�ւ̂ݕK�{�Ƃ���F�V�͖����\��������̂�.
    'If Index = 0 And 0 = dyn.RecordCount Then
    If 0 = dyn.RecordCount Then
        KeyCode = 0
        Call MsgBox("�Y���f�[�^�͑��݂��܂���.( " & fraOldNew(Index).Caption & "��" & lblBankcode(Index).Caption & ")", vbInformation, mCaption)
        Call txtDABANK(Index).SetFocus
        Exit Sub
    End If
    'txtDAYKED(Index).Number = 0
    If Not dyn.EOF Then
        If Index = 0 Then
            lblBankCount.Caption = dyn.RecordCount
            txtDAYKED(Index).Number = gdDBS.sysDate("YYYYMMDD")
'''2002/10/09 �z�X�g�f�[�^�̊֌W�Ńt�B�[���h���폜����
'''            lblGenzaiTekiyoBi.Caption = dyn.Fields("DAYKED")
        End If
        lblBankName(Index).Caption = dyn.Fields("DAKJNM")
    End If
End Sub
#End If

Private Sub txtDABANK_LostFocus(Index As Integer)
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim msg As String
        
    lblBankName(Index).Caption = ""
    txtDASITN(Index).Text = ""
    lblShitenName(Index).Caption = ""
    If "" = Trim(txtDABANK(Index).Text) Then
        Exit Sub
    End If
'''2002/10/09 �z�X�g�f�[�^�̊֌W�Ńt�B�[���h���폜����
'''    sql = "SELECT DAKJNM,DAYKED FROM tdBankMaster" & vbCrLf
    sql = "SELECT DAKJNM FROM tdBankMaster" & vbCrLf
    sql = sql & " WHERE DARKBN = '" & eBankRecordKubun.Bank & "'" & vbCrLf
    sql = sql & "   AND DABANK = '" & txtDABANK(Index).Text & "'" & vbCrLf
'''2002/10/09 �z�X�g�f�[�^�̊֌W�Ńt�B�[���h���폜����
'''    sql = sql & "   AND TO_CHAR(SYSDATE,'YYYYMMDD') BETWEEN DAYKST AND DAYKED" & vbCrLf
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    '//���E���Z�@�ւ̂ݕK�{�Ƃ���F�V�͖����\��������̂�.
    'If Index = 0 And 0 = dyn.RecordCount Then
    If 0 = dyn.RecordCount Then
        Call MsgBox("�Y���f�[�^�͑��݂��܂���.( " & fraOldNew(Index).Caption & "��" & lblBankcode(Index).Caption & ")", vbInformation, mCaption)
        Call txtDABANK(Index).SetFocus
        Exit Sub
    End If
    'txtDAYKED(Index).Number = 0
    If Not dyn.EOF Then
        If Index = 0 Then
            lblBankCount.Caption = dyn.RecordCount
            txtDAYKED(Index).Number = gdDBS.sysDate("YYYYMMDD")
'''2002/10/09 �z�X�g�f�[�^�̊֌W�Ńt�B�[���h���폜����
'''            lblGenzaiTekiyoBi.Caption = dyn.Fields("DAYKED")
        End If
        lblBankName(Index).Caption = dyn.Fields("DAKJNM")
    End If
End Sub

#If 0 Then
Private Sub txtDASITN_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not (KeyCode = vbKeyReturn) Then
        Exit Sub
    End If
    
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
        
    lblShitenName(Index).Caption = ""
    If "" = Trim(txtDASITN(Index).Text) Then
        Exit Sub
    End If
    sql = "SELECT DAKJNM FROM tdBankMaster" & vbCrLf
    sql = sql & " WHERE DARKBN = '" & eBankRecordKubun.Shiten & "'" & vbCrLf
    sql = sql & "   AND DABANK = '" & txtDABANK(Index).Text & "'" & vbCrLf
    sql = sql & "   AND DASITN = '" & txtDASITN(Index).Text & "'" & vbCrLf
'''2002/10/09 �z�X�g�f�[�^�̊֌W�Ńt�B�[���h���폜����
'''    sql = sql & "   AND TO_CHAR(SYSDATE,'YYYYMMDD') BETWEEN DAYKST AND DAYKED" & vbCrLf
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    '//���E���Z�@�ւ̂ݕK�{�Ƃ���F�V�͖����\��������̂�.
    'If Index = 0 And 0 = dyn.RecordCount Then
    If 0 = dyn.RecordCount Then
        KeyCode = 0
        Call MsgBox("�Y���f�[�^�͑��݂��܂���.( " & fraOldNew(Index).Caption & "��" & lblShitenCode(Index).Caption & ")", vbInformation, mCaption)
        Call txtDASITN(Index).SetFocus
        Exit Sub
    End If
    If Not dyn.EOF Then
        lblBankCount.Caption = dyn.RecordCount
        lblShitenName(Index).Caption = dyn.Fields("DAKJNM")
    End If
'    Call txtDAYKED(Index).SetFocus
End Sub
#End If

Private Sub txtDASITN_LostFocus(Index As Integer)
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
        
    lblShitenName(Index).Caption = ""
    If "" = Trim(txtDASITN(Index).Text) Then
        Exit Sub
    End If
    sql = "SELECT DAKJNM FROM tdBankMaster" & vbCrLf
    sql = sql & " WHERE DARKBN = '" & eBankRecordKubun.Shiten & "'" & vbCrLf
    sql = sql & "   AND DABANK = '" & txtDABANK(Index).Text & "'" & vbCrLf
    sql = sql & "   AND DASITN = '" & txtDASITN(Index).Text & "'" & vbCrLf
'''2002/10/09 �z�X�g�f�[�^�̊֌W�Ńt�B�[���h���폜����
'''    sql = sql & "   AND TO_CHAR(SYSDATE,'YYYYMMDD') BETWEEN DAYKST AND DAYKED" & vbCrLf
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    '//���E���Z�@�ւ̂ݕK�{�Ƃ���F�V�͖����\��������̂�.
    'If Index = 0 And 0 = dyn.RecordCount Then
    If 0 = dyn.RecordCount Then
        Call MsgBox("�Y���f�[�^�͑��݂��܂���.( " & fraOldNew(Index).Caption & "��" & lblShitenCode(Index).Caption & ")", vbInformation, mCaption)
        Call txtDASITN(Index).SetFocus
        Exit Sub
    End If
    If Not dyn.EOF Then
        lblBankCount.Caption = dyn.RecordCount
        lblShitenName(Index).Caption = dyn.Fields("DAKJNM")
    End If
'    Call txtDAYKED(Index).SetFocus
End Sub

Private Sub txtDAYKED_Change(Index As Integer)
    On Error GoTo txtDAYKED_ChangeError:
    '// Index = 0 �݂̂����͉\
    If Index = 0 Then
'        If txtDAYKED(eOldNew.OldBank).Year > 0 And txtDAYKED(eOldNew.OldBank).Month > 0 And txtDAYKED(eOldNew.OldBank).Day > 0 Then
        If txtDAYKED(eOldNew.OldBank).Number > 0 And Val(lblShoriKubun.Caption) = eShoriKubun.TouHaigou Then
            txtDAYKED(eOldNew.NewBank).Text = Format(DateSerial(txtDAYKED(eOldNew.OldBank).Year, txtDAYKED(eOldNew.OldBank).Month, txtDAYKED(eOldNew.OldBank).Day + 1), "YYYY/MM/DD")
        Else
            txtDAYKED(eOldNew.NewBank).Number = 0
        End If
    End If
    Exit Sub
txtDAYKED_ChangeError:
    Call MsgBox("���t�� " & Format(txtDAYKED(Index).MinDate, "yyyy/mm/dd") & " �ȏ�œ��͂��ĉ�����.", vbInformation, mCaption)
End Sub

Private Sub pMakeNewRecord()
    Dim sql As String
    Call gdDBS.Database.BeginTrans

'//2007/06/11 ��ʂ� AutoLog �ɂ������̂Ńg���K���~
'//      �����͐��䂵�Ȃ��F�ύX���ꂽ�Ƃ���ɓ��e���o��
'    Call gdDBS.TriggerControl("tcHogoshaMaster", False)

'''���Z�@�ւɑ΂��Ă͍X�V����K�v�Ȃ�
'''    '//��s�ɑ΂��Ă͓����E�ڍs���̃��R�[�h�ǉ������͖���
'''    '//�p�~��(�L���I����)���X�V
'''    sql = "UPDATE tdBankMaster SET "
'''    sql = sql & " DAYKED = " & txtDAYKED(eOldNew.OldBank).Number
'''    sql = sql & " WHERE DABANK = '" & txtDABANK(eOldNew.OldBank).Text & "'"
'''    If "" <> txtDASITN(eOldNew.OldBank).Text Then
'''        sql = sql & "   AND DASITN = '" & txtDASITN(eOldNew.OldBank).Text & "'"
'''    End If
'''    Call gdDBS.Database.ExecuteSQL(sql)

'///////////////////////////////////////////
'///////////////////////////////////////////
'///////////////////////////////////////////
'///////////////////////////////////////////
'///////////////////////////////////////////
    
    Call pMakeNewKeiyakusha     '//�_���
    Call pMakeNewHogosha        '//�ی��
    
    Call gdDBS.Database.CommitTrans

'//2007/06/11 ��ʂ� AutoLog �ɂ������̂Ńg���K���~
'//      �����͐��䂵�Ȃ��F�ύX���ꂽ�Ƃ���ɓ��e���o��
'    Call gdDBS.TriggerControl("tcHogoshaMaster", True)
    
    Call MsgBox("�����͐���I�����܂���.", vbInformation, mCaption)
pMakeNewRecordError:
    Call gdDBS.Database.Rollback

'//2007/06/11 ��ʂ� AutoLog �ɂ������̂Ńg���K���~
'//      �����͐��䂵�Ȃ��F�ύX���ꂽ�Ƃ���ɓ��e���o��
'    Call gdDBS.TriggerControl("tcHogoshaMaster", True)
    Call gdDBS.ErrorCheck(gdDBS.Database)

'// gdDBS.ErrorCheck() �̏�Ɉړ�
'//    Call gdDBS.Database.Rollback
End Sub

Private Sub pMakeNewKeiyakusha()
    '//�_��҃}�X�^���ǉ�
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim fld As Variant
    Dim flgUpdate As Boolean, flgInsert As Boolean
    Dim ix As Integer
    
    '//�_��҃}�X�^�e�[�u���̗񖼎擾
    fld = gdDBS.FieldNames("tbKeiyakushaMaster")
    
    sql = "SELECT * FROM tbKeiyakushaMaster"
    sql = sql & pMakeWhereSQL(eTable.Itakusha)
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//    sql = sql & " ORDER BY BAITKB,BAKYCD,BAKSCD"
    sql = sql & " ORDER BY BAITKB,BAKYCD"
'    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_DEFAULT)
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_DEFAULT)
    With dyn
        Do Until .EOF()
            flgUpdate = False
            flgInsert = False
            '//���ꂩ��쐬����f�[�^�� BASQNO=�{���Ƃ���̂�...�B
            '//�U���I�������K�p�I�����̃f�[�^�̂݁F�ȊO�͏C���ς݂̂͂��H
            '//BASQNO = �{���Ȃ��
            If .Fields("BASQNO") = gdDBS.sysDate("YYYYMMDD") Then
                '//�U���I����������t�Ȃ��
                If .Fields("BAFKED") > txtDAYKED(eOldNew.OldBank).Number Then
                    '//�U���I�������Z�b�g
                    flgUpdate = True
                    '//�V�K���R�[�h�͍쐬�ł��Ȃ�!!!
                    'flgInsert = lblShoriKubun.Caption <> eShoriKubun.Haishi
                End If
            Else
                '//�U���I����������t�Ȃ��
                If .Fields("BAFKED") > txtDAYKED(eOldNew.OldBank).Number Then
                    '//�U���I�������Z�b�g
                    flgUpdate = True
                    '//�p�~�łȂ���ΐV���Z�@�ւ��g�p�����_��҃f�[�^���쐬
                    flgInsert = lblShoriKubun.Caption <> eShoriKubun.Haishi
                End If
            End If
            '//���݂̃��R�[�h�̃R�s�[���쐬�F���Z�@�ւ͓���ւ�
            If flgInsert = True Then
                sql = "INSERT INTO tbKeiyakushaMaster("
                For ix = LBound(fld) To UBound(fld)
                    sql = sql & fld(ix) & ","
                Next ix
                sql = Left(sql, Len(sql) - 1) & ") SELECT " '�Ō�́u,�v���폜
                For ix = LBound(fld) To UBound(fld)
                    '//�ύX�l
                    Select Case fld(ix)
                    Case "BASQNO":  sql = sql & "TO_CHAR(SYSDATE,'YYYYMMDD'),"
                    Case "BABANK":  sql = sql & "'" & txtDABANK(eOldNew.NewBank).Text & "',"
                    Case "BASITN":
                        If "" <> txtDASITN(eOldNew.NewBank).Text Then
                            sql = sql & "'" & txtDASITN(eOldNew.NewBank).Text & "',"
                        Else
                            sql = sql & fld(ix) & ","
                        End If
                    Case "BAFKST":  sql = sql & txtDAYKED(eOldNew.NewBank).Number & ","
                    Case "BAFKED":  sql = sql & gdDBS.LastDay(0) & ","
                    Case "BAUSID":  sql = sql & gdDBS.ColumnDataSet(MainModule.gcBankBatchUpdateUser)
                    Case "BAUPDT":  sql = sql & "SYSDATE,"
                    Case Else:      sql = sql & fld(ix) & ","
                    End Select
                Next ix
                sql = Left(sql, Len(sql) - 1) '�Ō�́u,�v���폜
                sql = sql & " FROM tbKeiyakushaMaster"
                sql = sql & " WHERE BAITKB = '" & dyn.Fields("BAITKB") & "'"
                sql = sql & "   AND BAKYCD = '" & dyn.Fields("BAKYCD") & "'"
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//                sql = sql & "   AND BAKSCD = '" & dyn.Fields("BAKSCD") & "'"
                sql = sql & "   AND BASQNO = '" & dyn.Fields("BASQNO") & "'"
                Call gdDBS.Database.ExecuteSQL(sql)
            End If
            '//���݂̃��R�[�h��u����
            If flgUpdate = True Then
                Call .Edit
                .Fields("BAFKED").Value = txtDAYKED(eOldNew.OldBank).Number
                .Fields("BAUSID").Value = MainModule.gcBankBatchUpdateUser
                .Fields("BAUPDT").Value = gdDBS.sysDate
                Call .Update
            End If
            Call .MoveNext
        Loop
    End With
End Sub

Private Sub pMakeNewHogosha()
    '//�ی�҃}�X�^���ǉ�
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim fld As Variant
    Dim flgUpdate As Boolean, flgInsert As Boolean
    Dim ix As Integer
    
    '//�ی�҃}�X�^�e�[�u���̗񖼎擾
    fld = gdDBS.FieldNames("tcHogoshaMaster")
    
    sql = "SELECT * FROM tcHogoshaMaster"
    sql = sql & pMakeWhereSQL(eTable.Hogosha)
    sql = sql & " ORDER BY CAITKB,CAKYCD,CAKSCD,CAHGCD"
    Set dyn = gdDBS.OpenRecordset(sql)
    With dyn
        Do Until .EOF()
            flgUpdate = False
            flgInsert = False
            '//���ꂩ��쐬����f�[�^�� CASQNO=�{���Ƃ���̂�...�B
            '//�U�֏I�������K�p�I�����̃f�[�^�̂݁F�ȊO�͏C���ς݂̂͂��H
            '//CASQNO = �{���Ȃ��
            If .Fields("CASQNO") = gdDBS.sysDate("YYYYMMDD") Then
                '//�U�֏I����������t�Ȃ��
                If .Fields("CAFKED") > txtDAYKED(eOldNew.OldBank).Number Then
                    '//�U�֏I�������Z�b�g
                    flgUpdate = True
                    '//�V�K���R�[�h�͍쐬�ł��Ȃ�!!!
                    'flgInsert = lblShoriKubun.Caption <> eShoriKubun.Haishi
                End If
            Else
                '//�U�֏I����������t�Ȃ��
                If .Fields("CAFKED") > txtDAYKED(eOldNew.OldBank).Number Then
                    '//�U�֏I�������Z�b�g
                    flgUpdate = True
                    '//�p�~�łȂ���ΐV���Z�@�ւ��g�p�����ی�҃f�[�^���쐬
                    flgInsert = lblShoriKubun.Caption <> eShoriKubun.Haishi
                End If
            End If
            '//���݂̃��R�[�h�̃R�s�[���쐬�F���Z�@�ւ͓���ւ�
            If flgInsert = True Then
                sql = "INSERT INTO tcHogoshaMaster("
                For ix = LBound(fld) To UBound(fld)
                    sql = sql & fld(ix) & ","
                Next ix
                sql = Left(sql, Len(sql) - 1) & ") SELECT " '�Ō�́u,�v���폜
                For ix = LBound(fld) To UBound(fld)
                    '//�ύX�l
                    Select Case fld(ix)
                    Case "CASQNO":  sql = sql & "TO_CHAR(SYSDATE,'YYYYMMDD'),"
                    Case "CABANK":  sql = sql & "'" & txtDABANK(eOldNew.NewBank).Text & "',"
                    Case "CASITN":
                        If "" <> txtDASITN(eOldNew.NewBank).Text Then
                            sql = sql & "'" & txtDASITN(eOldNew.NewBank).Text & "',"
                        Else
                            sql = sql & fld(ix) & ","
                        End If
                    Case "CAFKST":  sql = sql & txtDAYKED(eOldNew.NewBank).Number & ","
                    Case "CAFKED":  sql = sql & gdDBS.LastDay(0) & ","
                    Case "CAUSID":  sql = sql & gdDBS.ColumnDataSet(MainModule.gcBankBatchUpdateUser)
                    Case "CAUPDT":  sql = sql & "SYSDATE,"
                    Case Else:      sql = sql & fld(ix) & ","
                    End Select
                Next ix
                sql = Left(sql, Len(sql) - 1) '�Ō�́u,�v���폜
                sql = sql & " FROM tcHogoshaMaster"
                sql = sql & " WHERE CAITKB = '" & dyn.Fields("CAITKB") & "'"
                sql = sql & "   AND CAKYCD = '" & dyn.Fields("CAKYCD") & "'"
                sql = sql & "   AND CAKSCD = '" & dyn.Fields("CAKSCD") & "'"
                sql = sql & "   AND CAHGCD = '" & dyn.Fields("CAHGCD") & "'"
                sql = sql & "   AND CASQNO = '" & dyn.Fields("CASQNO") & "'"
                Call gdDBS.Database.ExecuteSQL(sql)
            End If
            '//���݂̃��R�[�h��u����
            If flgUpdate = True Then
                Call .Edit
                .Fields("CAFKED").Value = txtDAYKED(eOldNew.OldBank).Number
                .Fields("CAUSID").Value = MainModule.gcBankBatchUpdateUser
                .Fields("CAUPDT").Value = gdDBS.sysDate
'//2006/04/26 �����R�[�h���V�K�����̂Ƃ� 1900/01/01 ����������F����ł���̂ɂ��܂ł��V�K�����̗l�ȐU�镑��������
                If IsNull(.Fields("CANWDT")) Then
                    .Fields("CANWDT").Value = "1900/01/01"
                End If
                Call .Update
            End If
            Call .MoveNext
        Loop
    End With
End Sub

Private Function pMakeWhereSQL(Optional ByVal vMode As Integer = -1) As String
    Dim sql As String
    Select Case vMode
    Case eTable.Itakusha
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//        sql = " WHERE (BAITKB,BAKYCD,BAKSCD,BASQNO) IN("
'//            sql = sql & " SELECT BAITKB,BAKYCD,BAKSCD,MAX(BASQNO) FROM tbKeiyakushaMaster"
        sql = " WHERE (BAITKB,BAKYCD,BASQNO) IN("
            sql = sql & " SELECT BAITKB,BAKYCD,MAX(BASQNO) FROM tbKeiyakushaMaster"
            '//�U�����Ԃ��L���ȃf�[�^
            sql = sql & " WHERE BAFKED  > " & txtDAYKED(eOldNew.OldBank).Number
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//            sql = sql & " GROUP BY BAITKB,BAKYCD,BAKSCD"
            sql = sql & " GROUP BY BAITKB,BAKYCD"
        sql = sql & ")"
        sql = sql & "   AND BAKKBN = '" & eBankKubun.KinnyuuKikan & "'"
        sql = sql & "   AND BABANK = '" & txtDABANK(eOldNew.OldBank).Text & "'"
        If "" <> txtDASITN(eOldNew.OldBank).Text Then
            sql = sql & "   AND BASITN = '" & txtDASITN(eOldNew.OldBank).Text & "'"
        End If
''        '//�U�����Ԃ��L���ȃf�[�^
''        sql = sql & "   AND TO_CHAR(SYSDATE,'YYYYMMDD') BETWEEN BAFKST AND BAFKED"
    Case eTable.Hogosha
        sql = sql & " WHERE (CAITKB,CAKYCD,CAKSCD,CAHGCD,CASQNO) IN("
            sql = sql & " SELECT CAITKB,CAKYCD,CAKSCD,CAHGCD,MAX(CASQNO) FROM tcHogoshaMaster"
            '//�U�����Ԃ��L���ȃf�[�^
            sql = sql & " WHERE CAFKED  > " & txtDAYKED(eOldNew.OldBank).Number
            sql = sql & " GROUP BY CAITKB,CAKYCD,CAKSCD,CAHGCD"
        sql = sql & ")"
        sql = sql & "   AND CAKKBN = '" & eBankKubun.KinnyuuKikan & "'"
        sql = sql & "   AND CABANK = '" & txtDABANK(eOldNew.OldBank).Text & "'"
        If "" <> txtDASITN(eOldNew.OldBank).Text Then
            sql = sql & "   AND CASITN = '" & txtDASITN(eOldNew.OldBank).Text & "'"
        End If
''        '//�U�֊��Ԃ��L���ȃf�[�^
''        sql = sql & "   AND TO_CHAR(SYSDATE,'YYYYMMDD') BETWEEN CAFKST AND CAFKED"
    End Select
    pMakeWhereSQL = sql
End Function

Private Sub mnuEnd_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

Private Sub txtDAYKED_DropOpen(Index As Integer, NoDefault As Boolean)
    txtDAYKED(Index).Calendar.Holidays = gdDBS.Holiday(txtDAYKED(Index).Year)
End Sub

