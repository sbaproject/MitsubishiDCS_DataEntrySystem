VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Begin VB.Form frmSystemInfomation 
   Caption         =   "��{���}�X�^�����e�i���X"
   ClientHeight    =   4845
   ClientLeft      =   2730
   ClientTop       =   2235
   ClientWidth     =   7665
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7665
   Begin imDate6Ctl.imDate txtAANXKZ 
      DataField       =   "AANXKZ"
      DataSource      =   "dbcSystem"
      Height          =   285
      Left            =   5220
      TabIndex        =   3
      Top             =   1560
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   503
      Calendar        =   "��{���}�X�^�����e�i���X.frx":0000
      Caption         =   "��{���}�X�^�����e�i���X.frx":0180
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "��{���}�X�^�����e�i���X.frx":01EC
      Keys            =   "��{���}�X�^�����e�i���X.frx":020A
      MouseIcon       =   "��{���}�X�^�����e�i���X.frx":0268
      Spin            =   "��{���}�X�^�����e�i���X.frx":0284
      AlignHorizontal =   0
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
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "    /  /  "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   -2
      CenturyMode     =   0
   End
   Begin imDate6Ctl.imDate txtAANXFK 
      DataField       =   "AANXFK"
      DataSource      =   "dbcSystem"
      Height          =   285
      Left            =   5220
      TabIndex        =   5
      Top             =   1980
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   503
      Calendar        =   "��{���}�X�^�����e�i���X.frx":02AC
      Caption         =   "��{���}�X�^�����e�i���X.frx":042C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "��{���}�X�^�����e�i���X.frx":0498
      Keys            =   "��{���}�X�^�����e�i���X.frx":04B6
      MouseIcon       =   "��{���}�X�^�����e�i���X.frx":0514
      Spin            =   "��{���}�X�^�����e�i���X.frx":0530
      AlignHorizontal =   0
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
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "    /  /  "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   -2
      CenturyMode     =   0
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "���~(&C)"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2700
      TabIndex        =   9
      Top             =   4080
      Width           =   1395
   End
   Begin imNumber6Ctl.imNumber txtAAFKDT 
      DataField       =   "AAFKDT"
      DataSource      =   "dbcSystem"
      Height          =   315
      Left            =   3360
      TabIndex        =   4
      Top             =   1980
      Width           =   375
      _Version        =   65537
      _ExtentX        =   661
      _ExtentY        =   556
      Calculator      =   "��{���}�X�^�����e�i���X.frx":0558
      Caption         =   "��{���}�X�^�����e�i���X.frx":0578
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "��{���}�X�^�����e�i���X.frx":05E4
      Keys            =   "��{���}�X�^�����e�i���X.frx":0602
      MouseIcon       =   "��{���}�X�^�����e�i���X.frx":064C
      Spin            =   "��{���}�X�^�����e�i���X.frx":0668
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#0"
      HighlightText   =   -1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   31
      MinValue        =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   5
      Value           =   1
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin imText6Ctl.imText ImText1 
      DataField       =   "AAYSNM"
      DataSource      =   "dbcSystem"
      Height          =   315
      Left            =   3360
      TabIndex        =   7
      Top             =   3300
      Width           =   2235
      _Version        =   65537
      _ExtentX        =   3942
      _ExtentY        =   556
      Caption         =   "��{���}�X�^�����e�i���X.frx":0690
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "��{���}�X�^�����e�i���X.frx":06FC
      Key             =   "��{���}�X�^�����e�i���X.frx":071A
      MouseIcon       =   "��{���}�X�^�����e�i���X.frx":075E
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
      MaxLength       =   15
      LengthAsByte    =   0
      Text            =   "ճ���ֳ �ַݷָ"
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
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�X�V(&U)"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   900
      TabIndex        =   8
      Top             =   4080
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "�I��(&X)"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5460
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
   Begin ORADCLibCtl.ORADC dbcSystem 
      Height          =   315
      Left            =   5580
      Top             =   2760
      Visible         =   0   'False
      Width           =   1875
      _Version        =   65536
      _ExtentX        =   3307
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
      RecordSource    =   "SELECT * FROM taSystemInformation a"
   End
   Begin imText6Ctl.imText ImText2 
      DataField       =   "AANAME"
      DataSource      =   "dbcSystem"
      Height          =   315
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   3675
      _Version        =   65537
      _ExtentX        =   6482
      _ExtentY        =   556
      Caption         =   "��{���}�X�^�����e�i���X.frx":077A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "��{���}�X�^�����e�i���X.frx":07E6
      Key             =   "��{���}�X�^�����e�i���X.frx":0804
      MouseIcon       =   "��{���}�X�^�����e�i���X.frx":0848
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
      MaxLength       =   50
      LengthAsByte    =   0
      Text            =   "�_�C�������h�t�@�N�^�[�@�������"
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
   Begin imText6Ctl.imText ImText3 
      DataField       =   "AAADDR"
      DataSource      =   "dbcSystem"
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   1140
      Width           =   3675
      _Version        =   65537
      _ExtentX        =   6482
      _ExtentY        =   556
      Caption         =   "��{���}�X�^�����e�i���X.frx":0864
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "��{���}�X�^�����e�i���X.frx":08D0
      Key             =   "��{���}�X�^�����e�i���X.frx":08EE
      MouseIcon       =   "��{���}�X�^�����e�i���X.frx":0932
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
      MaxLength       =   50
      LengthAsByte    =   0
      Text            =   "�s�d�k�@�O�R�|�R�Q�T�P�|�W�R�O�O"
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
   Begin imText6Ctl.imText ImText4 
      DataField       =   "AAYSNO"
      DataSource      =   "dbcSystem"
      Height          =   315
      Left            =   3360
      TabIndex        =   6
      Top             =   2880
      Width           =   495
      _Version        =   65537
      _ExtentX        =   873
      _ExtentY        =   556
      Caption         =   "��{���}�X�^�����e�i���X.frx":094E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "��{���}�X�^�����e�i���X.frx":09BA
      Key             =   "��{���}�X�^�����e�i���X.frx":09D8
      MouseIcon       =   "��{���}�X�^�����e�i���X.frx":0A1C
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
      LengthAsByte    =   0
      Text            =   "9900"
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
   Begin imNumber6Ctl.imNumber txtAAKZDT 
      DataField       =   "AAKZDT"
      DataSource      =   "dbcSystem"
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Top             =   1560
      Width           =   375
      _Version        =   65537
      _ExtentX        =   661
      _ExtentY        =   556
      Calculator      =   "��{���}�X�^�����e�i���X.frx":0A38
      Caption         =   "��{���}�X�^�����e�i���X.frx":0A58
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "��{���}�X�^�����e�i���X.frx":0AC4
      Keys            =   "��{���}�X�^�����e�i���X.frx":0AE2
      MouseIcon       =   "��{���}�X�^�����e�i���X.frx":0B2C
      Spin            =   "��{���}�X�^�����e�i���X.frx":0B48
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#0"
      HighlightText   =   -1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   31
      MinValue        =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   5
      Value           =   1
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin imText6Ctl.imText txtAANWDT 
      Height          =   315
      Left            =   3360
      TabIndex        =   23
      Top             =   2400
      Width           =   1875
      _Version        =   65537
      _ExtentX        =   3307
      _ExtentY        =   556
      Caption         =   "��{���}�X�^�����e�i���X.frx":0B70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "��{���}�X�^�����e�i���X.frx":0BDC
      Key             =   "��{���}�X�^�����e�i���X.frx":0BFA
      MouseIcon       =   "��{���}�X�^�����e�i���X.frx":0C3E
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
      Format          =   "H"
      FormatMode      =   0
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   20
      LengthAsByte    =   0
      Text            =   "2001/01/31 23:59:59"
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
   Begin VB.Label lblAANWDT 
      Alignment       =   1  '�E����
      BackColor       =   &H000000FF&
      Caption         =   "2003/01/30 22:10:09"
      DataField       =   "AANWDT"
      DataSource      =   "dbcSystem"
      Height          =   255
      Left            =   5340
      TabIndex        =   25
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   1  '�E����
      Caption         =   "�V�K�������"
      Height          =   255
      Left            =   1740
      TabIndex        =   24
      Top             =   2460
      Width           =   1515
   End
   Begin VB.Label Label10 
      Caption         =   "����U�֓�"
      Height          =   255
      Left            =   4200
      TabIndex        =   22
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "����U����"
      Height          =   195
      Left            =   4200
      TabIndex        =   21
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblSystemKey 
      Caption         =   "Label9"
      DataField       =   "AASKEY"
      DataSource      =   "dbcSystem"
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   19
      Top             =   60
      Width           =   1395
   End
   Begin VB.Label Label8 
      Caption         =   "��"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "��"
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   1620
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   1  '�E����
      Caption         =   "�X�֋ǖ���"
      Height          =   255
      Left            =   1740
      TabIndex        =   16
      Top             =   3360
      Width           =   1515
   End
   Begin VB.Label Label5 
      Alignment       =   1  '�E����
      Caption         =   "�X�֋ǔԍ�"
      Height          =   255
      Left            =   1740
      TabIndex        =   15
      Top             =   2940
      Width           =   1515
   End
   Begin VB.Label Label4 
      Alignment       =   1  '�E����
      Caption         =   "�����U�֊�� ����"
      Height          =   255
      Left            =   1380
      TabIndex        =   14
      Top             =   1620
      Width           =   1875
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      Caption         =   "�U����� ����"
      Height          =   255
      Left            =   1740
      TabIndex        =   13
      Top             =   2040
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���[��s��Џ��ݒn"
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
      Caption         =   "���[��s��Ж���"
      Height          =   255
      Left            =   1260
      TabIndex        =   11
      Top             =   780
      Width           =   1515
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
Attribute VB_Name = "frmSystemInfomation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm As New FormClass
Private mCaption As String

Private Sub pLockedControl(vMode As Boolean)
    cmdCancel.Enabled = vMode
    cmdUpdate.Enabled = vMode
End Sub

Private Sub pDatabaseRead()
    Dim sql As String
    sql = "SELECT * FROM taSystemInformation"
    sql = sql & " WHERE aaskey = '" & gdDBS.SystemKey & "'"
    dbcSystem.RecordSource = sql
    Call dbcSystem.Refresh
    If dbcSystem.Recordset.RecordCount = 0 Then
        Call dbcSystem.Recordset.AddNew
        'dbcSystem.Recordset.Fields("aaskey") = gdDBS.SystemKey
    Else
        Call dbcSystem.Recordset.MoveFirst
        Call dbcSystem.Recordset.Edit
    End If
    Call pLockedControl(False)
End Sub

Private Sub cmdCancel_Click()
    dbcSystem.UpdateControls
    Call pLockedControl(False)
    Call pDatabaseRead
End Sub

Private Sub cmdEnd_Click()
    '//��� Edit ��Ԃɂ���̂ŃL�����Z������B
    dbcSystem.UpdateControls
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    If lblSystemKey.Caption = "" Then
        lblSystemKey.Caption = gdDBS.SystemKey
    End If
    If "" = gdDBS.CheckDateType(txtAANWDT.Text) Then
        Call MsgBox("�V�K������������t�`���ł͂���܂���." & vbCrLf & vbCrLf & "�����FYYYY/MM/DD HH24:MI:SS", vbCritical + vbOKOnly, mCaption)
        Call lblAANWDT_Change
        Exit Sub
    End If
'//2007/02/05 UpdateRecord() �ł���ƃG���[���E���Ȃ��̂� Recordset.Update() �ł���悤�ɕύX
    On Error GoTo pUpdateRecordError
    lblAANWDT.Caption = txtAANWDT.Text
'//2007/02/05 UpdateRecord() �ł���ƃG���[���E���Ȃ��̂� Recordset.Update() �ł���悤�ɕύX
'//    dbcSystem.UpdateRecord
    dbcSystem.Recordset.Update
    Call pLockedControl(False)
    Call pDatabaseRead
    Exit Sub
pUpdateRecordError:
    Call MsgBox("�X�V�������ɃG���[���������܂���." & vbCrLf & vbCrLf & Error, vbCritical + vbOKOnly, mCaption)
End Sub

Private Sub Form_Activate()
    Call pLockedControl(False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '//�ύX���ꂽ�H EditMode = ���ݍs�̌��݂̕ҏW��Ԃ�߂��܂�
'    Call pLockedControl(dbcSystem.EditMode <> editOption.ORADATA_EDITNONE)
    Call pLockedControl(dbcSystem.EditMode <> OracleConstantModule.ORADATA_EDITNONE)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    Call pDatabaseRead
'''pDatabaseRead() ���ł��Ă���
'''    dbcSystem.UpdateControls
'''    Call pLockedControl(False)
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSystemInfomation = Nothing
    Set mForm = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub lblAANWDT_Change()
    txtAANWDT.Text = Format(CVDate(lblAANWDT.Caption), "YYYY/MM/DD hh:nn:ss")
End Sub

Private Sub txtAAFKDT_InvalidRange(Restore As Boolean)
    Call MsgBox(txtAAFKDT.MinValue & "�`" & txtAAFKDT.MaxValue & "�͈̔͂œ��͂��ĉ�����.", vbInformation + vbOKOnly, mCaption)
    Call txtAAFKDT.SetFocus
End Sub

Private Sub txtAAKZDT_InvalidRange(Restore As Boolean)
    Call MsgBox(txtAAFKDT.MinValue & "�`" & txtAAFKDT.MaxValue & "�͈̔͂œ��͂��ĉ�����.", vbInformation + vbOKOnly, mCaption)
    Call txtAAKZDT.SetFocus
End Sub

Private Sub txtAANWDT_Change()
    Call pLockedControl(True)
End Sub

Private Sub txtAANXFK_Change()
    Call pLockedControl(True)
End Sub

Private Sub txtAANXFK_DropOpen(NoDefault As Boolean)
    txtAANXFK.Calendar.Holidays = gdDBS.Holiday(txtAANXFK.Year)
End Sub

Private Sub txtAANXKZ_Change()
    Call pLockedControl(True)
End Sub

Private Sub txtAANXKZ_DropOpen(NoDefault As Boolean)
    txtAANXKZ.Calendar.Holidays = gdDBS.Holiday(txtAANXKZ.Year)
End Sub

Private Sub mnuEnd_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

