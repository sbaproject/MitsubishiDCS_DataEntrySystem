VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHogoshaMaster 
   Caption         =   "�ی�҃}�X�^�����e�i���X"
   ClientHeight    =   7785
   ClientLeft      =   1815
   ClientTop       =   2175
   ClientWidth     =   10125
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
   ScaleHeight     =   7785
   ScaleWidth      =   10125
   Begin MSComCtl2.UpDown spnRireki 
      Height          =   420
      Left            =   2760
      TabIndex        =   12
      ToolTipText     =   "�O��̗����Ɉړ�"
      Top             =   1800
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   741
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.ComboBox cboABKJNM 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "�ی�҃}�X�^�����e�i���X.frx":0000
      Left            =   1800
      List            =   "�ی�҃}�X�^�����e�i���X.frx":000D
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Tag             =   "InputKey"
      Top             =   900
      Width           =   1755
   End
   Begin VB.Frame fraKinnyuuKikan 
      Caption         =   "�U�֌���"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   5040
      TabIndex        =   9
      Top             =   300
      Width           =   4635
      Begin VB.Frame fraBank 
         BackColor       =   &H00FF8080&
         Caption         =   "�X�֋�"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   1020
         Width           =   4035
         Begin imText6Ctl.imText txtCAYBTK 
            DataField       =   "CAYBTK"
            DataSource      =   "dbcHogoshaMaster"
            Height          =   285
            Left            =   1860
            TabIndex        =   25
            Top             =   480
            Width           =   375
            _Version        =   65537
            _ExtentX        =   661
            _ExtentY        =   503
            Caption         =   "�ی�҃}�X�^�����e�i���X.frx":002B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":0097
            Key             =   "�ی�҃}�X�^�����e�i���X.frx":00B5
            MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":00F9
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
            AllowSpace      =   0
            Format          =   "9"
            FormatMode      =   0
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   3
            LengthAsByte    =   0
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
         Begin imText6Ctl.imText txtCAYBTN 
            DataField       =   "CAYBTN"
            DataSource      =   "dbcHogoshaMaster"
            Height          =   285
            Left            =   1860
            TabIndex        =   26
            Top             =   960
            Width           =   855
            _Version        =   65537
            _ExtentX        =   1508
            _ExtentY        =   503
            Caption         =   "�ی�҃}�X�^�����e�i���X.frx":0115
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":0181
            Key             =   "�ی�҃}�X�^�����e�i���X.frx":019F
            MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":01E3
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
            AllowSpace      =   0
            Format          =   "9"
            FormatMode      =   0
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   8
            LengthAsByte    =   0
            Text            =   "12345678"
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
         Begin VB.Label lblTsuchoBango 
            Alignment       =   1  '�E����
            Caption         =   "�ʒ��ԍ�"
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
            Left            =   360
            TabIndex        =   53
            Top             =   960
            Width           =   1275
         End
         Begin VB.Label Label23 
            Alignment       =   1  '�E����
            Caption         =   "�ʒ��L��"
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
            Left            =   360
            TabIndex        =   52
            Top             =   480
            Width           =   1275
         End
      End
      Begin VB.Frame fraBank 
         BackColor       =   &H00FFFF00&
         Caption         =   "���ԋ��Z�@��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   0
         Left            =   480
         TabIndex        =   17
         Top             =   420
         Width           =   3855
         Begin imText6Ctl.imText txtCAKZNO 
            DataField       =   "CAKZNO"
            DataSource      =   "dbcHogoshaMaster"
            Height          =   285
            Left            =   1140
            TabIndex        =   23
            Top             =   1380
            Width           =   795
            _Version        =   65537
            _ExtentX        =   1402
            _ExtentY        =   503
            Caption         =   "�ی�҃}�X�^�����e�i���X.frx":01FF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":026B
            Key             =   "�ی�҃}�X�^�����e�i���X.frx":0289
            MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":02CD
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
         Begin imText6Ctl.imText txtCASITN 
            DataField       =   "CASITN"
            DataSource      =   "dbcHogoshaMaster"
            Height          =   285
            Left            =   1200
            TabIndex        =   19
            Top             =   660
            Width           =   375
            _Version        =   65537
            _ExtentX        =   661
            _ExtentY        =   503
            Caption         =   "�ی�҃}�X�^�����e�i���X.frx":02E9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":0355
            Key             =   "�ی�҃}�X�^�����e�i���X.frx":0373
            MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":03B7
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
            AllowSpace      =   0
            Format          =   "9"
            FormatMode      =   0
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   3
            LengthAsByte    =   0
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
         Begin imText6Ctl.imText txtCABANK 
            DataField       =   "CABANK"
            DataSource      =   "dbcHogoshaMaster"
            Height          =   285
            Left            =   1200
            TabIndex        =   18
            Top             =   300
            Width           =   495
            _Version        =   65537
            _ExtentX        =   873
            _ExtentY        =   503
            Caption         =   "�ی�҃}�X�^�����e�i���X.frx":03D3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":043F
            Key             =   "�ی�҃}�X�^�����e�i���X.frx":045D
            MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":04A1
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
            AllowSpace      =   0
            Format          =   "9"
            FormatMode      =   0
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   4
            LengthAsByte    =   0
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
         Begin VB.Frame fraKouzaShubetsu 
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  '�Ȃ�
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   1140
            TabIndex        =   54
            Top             =   900
            Width           =   2535
            Begin VB.OptionButton optCAKZSB 
               Caption         =   "����"
               BeginProperty Font 
                  Name            =   "�l�r �o�S�V�b�N"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   840
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   180
               Width           =   675
            End
            Begin VB.OptionButton optCAKZSB 
               Caption         =   "����"
               BeginProperty Font 
                  Name            =   "�l�r �o�S�V�b�N"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   180
               Width           =   675
            End
            Begin VB.OptionButton optCAKZSB 
               BackColor       =   &H000000FF&
               Caption         =   "Dummy"
               BeginProperty Font 
                  Name            =   "�l�r �o�S�V�b�N"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   1500
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   480
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label lblCAKZSB 
               BackColor       =   &H000000FF&
               Caption         =   "�������"
               DataField       =   "CAKZSB"
               DataSource      =   "dbcHogoshaMaster"
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
               Left            =   1620
               TabIndex        =   55
               Top             =   180
               Width           =   795
            End
         End
         Begin VB.Label lblBankName 
            Caption         =   "�����O�H�T�U�Vx"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   61
            Top             =   300
            Width           =   1935
         End
         Begin VB.Label Label12 
            Caption         =   "�����s"
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
            Left            =   240
            TabIndex        =   60
            Top             =   300
            Width           =   795
         End
         Begin VB.Label Label13 
            Caption         =   "����x�X"
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
            Left            =   240
            TabIndex        =   59
            Top             =   660
            Width           =   795
         End
         Begin VB.Label Label14 
            Caption         =   "�������"
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
            Left            =   240
            TabIndex        =   58
            Top             =   1020
            Width           =   795
         End
         Begin VB.Label Label15 
            Caption         =   "�����ԍ�"
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
            Left            =   240
            TabIndex        =   57
            Top             =   1380
            Width           =   795
         End
         Begin VB.Label lblShitenName 
            Caption         =   "���R�S�T�U�Vx"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   56
            Top             =   660
            Width           =   1935
         End
      End
      Begin VB.OptionButton optCAKKBN 
         Caption         =   "�X�֋�"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2100
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optCAKKBN 
         Caption         =   "���ԋ��Z�@��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   1395
      End
      Begin imText6Ctl.imText txtCAKZNM 
         DataField       =   "CAKZNM"
         DataSource      =   "dbcHogoshaMaster"
         Height          =   285
         Left            =   420
         TabIndex        =   27
         Top             =   2580
         Width           =   3735
         _Version        =   65537
         _ExtentX        =   6588
         _ExtentY        =   503
         Caption         =   "�ی�҃}�X�^�����e�i���X.frx":04BD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":0529
         Key             =   "�ی�҃}�X�^�����e�i���X.frx":0547
         MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":058B
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
         MaxLength       =   40
         LengthAsByte    =   0
         Text            =   "����Ҳ����Ҳ...........................*"
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
      Begin VB.Label Label28 
         Alignment       =   1  '�E����
         Caption         =   "�������`�l(�J�i)"
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
         Left            =   480
         TabIndex        =   76
         Top             =   2340
         Width           =   1395
      End
      Begin VB.Label lblCAKKBN 
         BackColor       =   &H000000FF&
         Caption         =   "���Z�@�֎��"
         DataField       =   "CAKKBN"
         DataSource      =   "dbcHogoshaMaster"
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
         Left            =   3180
         TabIndex        =   62
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkCAKYFG 
      Caption         =   "���"
      DataField       =   "CAKYFG"
      Height          =   315
      Left            =   4200
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3540
      Width           =   675
   End
   Begin VB.Frame fraBankList 
      Caption         =   "���Z�@�փ��X�g"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   4920
      TabIndex        =   31
      Top             =   3300
      Width           =   4875
      Begin VB.ComboBox cboBankYomi 
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "�ی�҃}�X�^�����e�i���X.frx":05A7
         Left            =   1500
         List            =   "�ی�҃}�X�^�����e�i���X.frx":05CC
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   180
         Width           =   855
      End
      Begin VB.ComboBox cboShitenYomi 
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "�ی�҃}�X�^�����e�i���X.frx":060E
         Left            =   3900
         List            =   "�ی�҃}�X�^�����e�i���X.frx":0633
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdKakutei 
         Caption         =   "�m��(&K)"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3660
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2700
         Width           =   975
      End
      Begin ORADCLibCtl.ORADC dbcShiten 
         Height          =   315
         Left            =   1920
         Top             =   2640
         Visible         =   0   'False
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   207
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.01
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DatabaseName    =   "dcssvr03"
         Connect         =   "kumon/kumon"
         RecordSource    =   ""
      End
      Begin ORADCLibCtl.ORADC dbcBank 
         Height          =   315
         Left            =   180
         Top             =   2640
         Visible         =   0   'False
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   207
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.01
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DatabaseName    =   "dcssvr03"
         Connect         =   "kumon/kumon"
         RecordSource    =   ""
      End
      Begin MSDBCtls.DBList dblBankList 
         Bindings        =   "�ی�҃}�X�^�����e�i���X.frx":0675
         Height          =   2040
         Left            =   120
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   540
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   3598
         _Version        =   393216
         IntegralHeight  =   0   'False
         ListField       =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDBCtls.DBList dblShitenList 
         Bindings        =   "�ی�҃}�X�^�����e�i���X.frx":068B
         Height          =   2040
         Left            =   2400
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   540
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   3598
         _Version        =   393216
         IntegralHeight  =   0   'False
         ListField       =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label24 
         Caption         =   "���Z�@�� �ǂ݁�"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label25 
         Caption         =   "�x�X�@�@�@�@�ǂ݁�"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2460
         TabIndex        =   63
         Top             =   240
         Width           =   1395
      End
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
      Left            =   2400
      TabIndex        =   29
      Top             =   6720
      Width           =   1395
   End
   Begin VB.Frame fraUpdateKubun 
      Caption         =   "�����敪"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   32
      Tag             =   "InputKey"
      Top             =   120
      Width           =   3735
      Begin VB.OptionButton optShoriKubun 
         Caption         =   "�Q��"
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
         Index           =   3
         Left            =   2880
         TabIndex        =   78
         Tag             =   "InputKey"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optShoriKubun 
         Caption         =   "�C��"
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
         Index           =   1
         Left            =   1140
         TabIndex        =   36
         Tag             =   "InputKey"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optShoriKubun 
         Caption         =   "�폜"
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
         Index           =   2
         Left            =   2040
         TabIndex        =   38
         Tag             =   "InputKey"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optShoriKubun 
         Caption         =   "�V�K"
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
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Tag             =   "InputKey"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblShoriKubun 
         BackColor       =   &H000000FF&
         Caption         =   "�����敪"
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
         Left            =   1380
         TabIndex        =   50
         Top             =   120
         Width           =   975
      End
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
      Left            =   660
      TabIndex        =   28
      Top             =   6720
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
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
      Left            =   8100
      TabIndex        =   30
      Top             =   6720
      Width           =   1395
   End
   Begin ORADCLibCtl.ORADC dbcHogoshaMaster 
      Height          =   315
      Left            =   5760
      Top             =   7080
      Visible         =   0   'False
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
      _ExtentY        =   556
      _StockProps     =   207
      BackColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.01
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   "srv0905"
      Connect         =   "wao2/wao2"
      RecordSource    =   "SELECT * FROM tcHogoshaMaster"
   End
   Begin imDate6Ctl.imDate txtCAKYxx 
      DataField       =   "CAKYST"
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   15
      Top             =   5700
      Visible         =   0   'False
      Width           =   795
      _Version        =   65537
      _ExtentX        =   1402
      _ExtentY        =   556
      Calendar        =   "�ی�҃}�X�^�����e�i���X.frx":06A3
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":0823
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":088F
      Keys            =   "�ی�҃}�X�^�����e�i���X.frx":08AD
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":090B
      Spin            =   "�ی�҃}�X�^�����e�i���X.frx":0927
      AlignHorizontal =   2
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   255
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   1
      DisplayFormat   =   "yyyy/mm"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy/mm"
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
      Text            =   "    /  "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   -2
      CenturyMode     =   0
   End
   Begin imDate6Ctl.imDate txtCAKYxx 
      DataField       =   "CAKYED"
      Height          =   315
      Index           =   1
      Left            =   3000
      TabIndex        =   16
      Top             =   5700
      Visible         =   0   'False
      Width           =   795
      _Version        =   65537
      _ExtentX        =   1402
      _ExtentY        =   556
      Calendar        =   "�ی�҃}�X�^�����e�i���X.frx":094F
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":0ACF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":0B3B
      Keys            =   "�ی�҃}�X�^�����e�i���X.frx":0B59
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":0BB7
      Spin            =   "�ی�҃}�X�^�����e�i���X.frx":0BD3
      AlignHorizontal =   2
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   255
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   1
      DisplayFormat   =   "yyyy/mm"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy/mm"
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
      Text            =   "    /  "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   -2
      CenturyMode     =   0
   End
   Begin imDate6Ctl.imDate txtCAFKxx 
      DataField       =   "CAFKST"
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   6
      Top             =   3540
      Width           =   795
      _Version        =   65537
      _ExtentX        =   1402
      _ExtentY        =   556
      Calendar        =   "�ی�҃}�X�^�����e�i���X.frx":0BFB
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":0D7B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":0DE7
      Keys            =   "�ی�҃}�X�^�����e�i���X.frx":0E05
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":0E63
      Spin            =   "�ی�҃}�X�^�����e�i���X.frx":0E7F
      AlignHorizontal =   2
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   1
      DisplayFormat   =   "yyyy/mm"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy/mm"
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
      Text            =   "    /  "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   -2
      CenturyMode     =   0
   End
   Begin imText6Ctl.imText txtCAKJNM 
      DataField       =   "CAKJNM"
      DataSource      =   "dbcHogoshaMaster"
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   2220
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":0EA7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":0F13
      Key             =   "�ی�҃}�X�^�����e�i���X.frx":0F31
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":0F75
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
      MaxLength       =   30
      LengthAsByte    =   -1
      Text            =   "���������D�D�D�D�D�D�D�D�D�D��"
      Furigana        =   -1
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
   Begin imText6Ctl.imText txtCAKYCD 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "InputKey"
      Top             =   1320
      Width           =   795
      _Version        =   65537
      _ExtentX        =   1402
      _ExtentY        =   503
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":0F91
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":0FFD
      Key             =   "�ی�҃}�X�^�����e�i���X.frx":101B
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":105F
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
   Begin imText6Ctl.imText txtCAHGCD 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Tag             =   "InputKey"
      Top             =   1860
      Width           =   855
      _Version        =   65537
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":107B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":10E7
      Key             =   "�ی�҃}�X�^�����e�i���X.frx":1105
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":1149
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
      AllowSpace      =   0
      Format          =   "9"
      FormatMode      =   0
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   8
      LengthAsByte    =   0
      Text            =   "12345678"
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
   Begin ORADCLibCtl.ORADC dbcItakushaMaster 
      Height          =   315
      Left            =   5760
      Top             =   6660
      Visible         =   0   'False
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
      _ExtentY        =   556
      _StockProps     =   207
      Caption         =   "taItakushaMaster"
      BackColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.01
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   "dcssvr03"
      Connect         =   "kumon/kumon"
      RecordSource    =   "SELECT * FROM taItakushaMaster"
   End
   Begin imText6Ctl.imText ImText1 
      DataField       =   "CASTNM"
      DataSource      =   "dbcHogoshaMaster"
      Height          =   465
      Left            =   1800
      TabIndex        =   5
      Top             =   2940
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   820
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":1165
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":11D1
      Key             =   "�ی�҃}�X�^�����e�i���X.frx":11EF
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":1233
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
      MultiLine       =   -1
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   50
      LengthAsByte    =   -1
      Text            =   "���k�����D�D�D�D�D�D�D�D�D�D��"
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
   Begin imText6Ctl.imText txtCAKNNM 
      DataField       =   "CAKNNM"
      DataSource      =   "dbcHogoshaMaster"
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   2580
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":124F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":12BB
      Key             =   "�ی�҃}�X�^�����e�i���X.frx":12D9
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":131D
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
      MaxLength       =   30
      LengthAsByte    =   -1
      Text            =   "�żҲ........................*"
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
   Begin imDate6Ctl.imDate txtCAFKxx 
      DataField       =   "CAFKED"
      Height          =   315
      Index           =   1
      Left            =   3120
      TabIndex        =   7
      Top             =   3540
      Width           =   795
      _Version        =   65537
      _ExtentX        =   1402
      _ExtentY        =   556
      Calendar        =   "�ی�҃}�X�^�����e�i���X.frx":1339
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":14B9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":1525
      Keys            =   "�ی�҃}�X�^�����e�i���X.frx":1543
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":15A1
      Spin            =   "�ی�҃}�X�^�����e�i���X.frx":15BD
      AlignHorizontal =   2
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   1
      DisplayFormat   =   "yyyy/mm"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy/mm"
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
      Text            =   "    /  "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   -2
      CenturyMode     =   0
   End
   Begin VB.Label lblSaveFKST 
      BackColor       =   &H000000FF&
      Caption         =   "�U�֊J�n��"
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
      Left            =   1875
      TabIndex        =   80
      Top             =   4350
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblCAKYDT 
      BackColor       =   &H000000FF&
      Caption         =   "����"
      DataField       =   "CAKYDT"
      DataSource      =   "dbcHogoshaMaster"
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
      Left            =   4050
      TabIndex        =   79
      Top             =   4275
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label4 
      Alignment       =   1  '�E����
      Caption         =   "�ی�Җ�(�J�i)"
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
      Left            =   300
      TabIndex        =   77
      Top             =   2580
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      Caption         =   "���k����"
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
      Left            =   240
      TabIndex        =   75
      Top             =   2985
      Width           =   1395
   End
   Begin VB.Label lblCAADDT 
      BackColor       =   &H000000FF&
      Caption         =   "�쐬��"
      DataField       =   "CAADDT"
      DataSource      =   "dbcHogoshaMaster"
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
      Left            =   7560
      TabIndex        =   74
      Top             =   6900
      Width           =   1755
   End
   Begin VB.Label lblCAKYFG 
      BackColor       =   &H000000FF&
      Caption         =   "���t���O"
      DataField       =   "CAKYFG"
      DataSource      =   "dbcHogoshaMaster"
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
      Left            =   4065
      TabIndex        =   73
      Top             =   3960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblCAFKxx 
      BackColor       =   &H000000FF&
      Caption         =   "�U�֏I����"
      DataField       =   "CAFKED"
      DataSource      =   "dbcHogoshaMaster"
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
      Index           =   1
      Left            =   2940
      TabIndex        =   72
      Top             =   4020
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblCAFKxx 
      BackColor       =   &H000000FF&
      Caption         =   "�U�֊J�n��"
      DataField       =   "CAFKST"
      DataSource      =   "dbcHogoshaMaster"
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
      Index           =   0
      Left            =   1920
      TabIndex        =   71
      Top             =   4020
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblCAKYxx 
      BackColor       =   &H000000FF&
      Caption         =   "�_��I����"
      DataField       =   "CAKYED"
      DataSource      =   "dbcHogoshaMaster"
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
      Index           =   1
      Left            =   2820
      TabIndex        =   70
      Top             =   6180
      Width           =   975
   End
   Begin VB.Label lblCAKYxx 
      BackColor       =   &H000000FF&
      Caption         =   "�_��J�n��"
      DataField       =   "CAKYST"
      DataSource      =   "dbcHogoshaMaster"
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
      Index           =   0
      Left            =   1800
      TabIndex        =   69
      Top             =   6180
      Width           =   975
   End
   Begin VB.Label lblCAUSID 
      BackColor       =   &H000000FF&
      Caption         =   "�X�V��"
      DataField       =   "CAUSID"
      DataSource      =   "dbcHogoshaMaster"
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
      Left            =   7560
      TabIndex        =   68
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label lblCAUPDT 
      BackColor       =   &H000000FF&
      Caption         =   "�X�V��"
      DataField       =   "CAUPDT"
      DataSource      =   "dbcHogoshaMaster"
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
      Left            =   7560
      TabIndex        =   67
      Top             =   7200
      Width           =   1755
   End
   Begin VB.Label lblCAITKB 
      BackColor       =   &H000000FF&
      Caption         =   "�ϑ��ҋ敪"
      DataField       =   "CAITKB"
      DataSource      =   "dbcHogoshaMaster"
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
      Left            =   3780
      TabIndex        =   66
      Top             =   660
      Width           =   975
   End
   Begin VB.Label Label26 
      Alignment       =   1  '�E����
      Caption         =   "�ϑ��ҋ敪"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   65
      Tag             =   "InputKey"
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label lblCASQNO 
      BackColor       =   &H000000FF&
      Caption         =   "�ی�҂r�d�p"
      DataField       =   "CASQNO"
      DataSource      =   "dbcHogoshaMaster"
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
      Left            =   3960
      TabIndex        =   14
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblCAKYCD 
      BackColor       =   &H000000FF&
      Caption         =   "�_��Ҕԍ�"
      DataField       =   "CAKYCD"
      DataSource      =   "dbcHogoshaMaster"
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
      Left            =   3780
      TabIndex        =   51
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label lblCAHGCD 
      BackColor       =   &H000000FF&
      Caption         =   "�ی�Ҕԍ�"
      DataField       =   "CAHGCD"
      DataSource      =   "dbcHogoshaMaster"
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
      Left            =   3960
      TabIndex        =   13
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label19"
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
      Left            =   8400
      TabIndex        =   49
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label lblBAKJNM 
      Caption         =   "�c���@�r�F"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2760
      TabIndex        =   48
      Tag             =   "InputKey"
      Top             =   1380
      Width           =   2235
   End
   Begin VB.Label Label10 
      Alignment       =   1  '�E����
      Caption         =   "�`"
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
      Left            =   2700
      TabIndex        =   47
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label7 
      Alignment       =   1  '�E����
      BackColor       =   &H000000FF&
      Caption         =   "�`"
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
      Left            =   2700
      TabIndex        =   46
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblKeiyakushaCode 
      Alignment       =   1  '�E����
      Caption         =   "�I�[�i�[�ԍ�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   45
      Tag             =   "InputKey"
      Top             =   1380
      Width           =   1275
   End
   Begin VB.Label lblHogoshaCode 
      Alignment       =   1  '�E����
      Caption         =   "�ی�Ҕԍ�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   44
      Tag             =   "InputKey"
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
      Caption         =   "�ی�Җ�(����)"
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
      Left            =   285
      TabIndex        =   43
      Top             =   2265
      Width           =   1395
   End
   Begin VB.Label Label18 
      Alignment       =   1  '�E����
      Caption         =   "�����U�֊���"
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
      Left            =   360
      TabIndex        =   42
      Top             =   3600
      Width           =   1275
   End
   Begin VB.Label Label16 
      Alignment       =   1  '�E����
      BackColor       =   &H000000FF&
      Caption         =   "�_�����"
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
      Left            =   480
      TabIndex        =   41
      Top             =   5760
      Visible         =   0   'False
      Width           =   1275
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
Attribute VB_Name = "frmHogoshaMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm As New FormClass
Private mCaption As String
Private mBankChange As Boolean  '//2006/08/22 ???_Change �C�x���g����s=>�x�X�ɋ�������
'//2013/02/26 �����ύX���̍X�V���̒ǉ��X�V�̍ۂɂQ�x pUpdateRecord() �����s�����̂𐧌䂷��
Private mRirekiAddNewUpdate As Boolean

Private Sub pLockedControl(blMode As Boolean)
    Call mForm.LockedControl(blMode)
'    cboBankYomi.ListIndex = -1
'    dblBankList.ListField = ""
'    dblBankList.Refresh
    Call dblBankList.ReFill
    '//dblBankList.Refresh() �����s����Ɖ��͕s�v
'    cboShitenYomi.ListIndex = -1
'    dblShitenList.ListField = ""
'    dblShitenList.Refresh
    Call dblShitenList.ReFill
    cmdEnd.Enabled = blMode
    spnRireki.Visible = False
End Sub

Private Sub chkCAKYFG_Click()
    lblCAKYFG.Caption = chkCAKYFG.Value
    If chkCAKYFG.Value = 0 Then
        lblCAKYDT.Caption = "20991231"
    Else
        lblCAKYDT.Caption = Format(Now(), "yyyyMMdd")
    End If
End Sub

Private Sub chkCAKYFG_KeyDown(KeyCode As Integer, Shift As Integer)
    '//���t���O��ݒ肵���̂ŏI�����̓��͂𑣂�.
    '//KeyCode & Shift ���N���A���Ȃ��ƃo�b�t�@�Ɏc��H
    KeyCode = 0
    Shift = 0
    chkCAKYFG.Value = Choose(chkCAKYFG.Value + 1, 1, 0, 0)  '// Index=1,2,3
    'Call MsgBox("���̕ύX�����m���܂����B" & vbCrLf & vbCrLf & "�����U�֊��� �I�����̍Đݒ�����ĉ�����.", vbInformation + vbOKOnly, mCaption)
    Call txtCAFKxx(1).SetFocus
End Sub

Private Sub chkCAKYFG_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '//���t���O��ݒ肵���̂ŏI�����̓��͂𑣂�.
    If Button = vbLeftButton Then
        Call chkCAKYFG_KeyDown(vbKeySpace, 0)
    End If
End Sub

Private Sub lblCAKYFG_Change()
    chkCAKYFG.Value = Val(lblCAKYFG.Caption)
End Sub

Private Function pUpdateRecord() As Boolean
#If D20060424 Then
'///////////////////////////////////////////////////////////////////////////////////////////
'//2006/04/24 ��������F�����ԍ��̃��j�[�N�����`�F�b�N�F�����ԍ��͂Ȃ����j�[�N����O�������H
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    sql = "SELECT * FROM tcHogoshaMaster"
    sql = sql & " WHERE CAITKB = '" & lblCAITKB.Caption & "'"
    sql = sql & "   AND CAKYCD = '" & lblCAKYCD.Caption & "'"
    sql = sql & "   AND CAHGCD = '" & lblCAHGCD.Caption & "'"
    sql = sql & "   AND CASQNO =  " & lblCASQNO.Caption
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If Not dyn.EOF Then
        If optShoriKubun(eShoriKubun.Add).Value = True Then
            Call MsgBox("���Ƀf�[�^�����݂��܂�.(" & lblHogoshaCode.Caption & ")", vbCritical, mCaption)
            Exit Function
        End If
    End If
'//2006/04/24 �����܂ŁF�����ԍ��̃��j�[�N�����`�F�b�N�F�����ԍ��͂Ȃ����j�[�N����O�������H
'///////////////////////////////////////////////////////////////////////////////////////////
#End If

'''//2002/10/18 ���̂܂܂̓��t�Ƃ���
'''    lblCAKYxx(0).Caption = gdDBS.FirstDay(txtCAKYxx(0).Number)
'''    lblCAKYxx(1).Caption = gdDBS.LastDay(txtCAKYxx(1).Number)
'''    lblCAFKxx(0).Caption = gdDBS.FirstDay(txtCAFKxx(0).Number)
'''    lblCAFKxx(1).Caption = gdDBS.LastDay(txtCAFKxx(1).Number)
'//2007/02/05 UpdateRecord() �ł���ƃG���[���E���Ȃ��̂� Recordset.Update() �ł���悤�ɕύX
    On Error GoTo pUpdateRecordError
    lblCAFKxx(0).Caption = Val(gdDBS.Nz(txtCAFKxx(0).Number))
    lblCAFKxx(1).Caption = Val(gdDBS.Nz(txtCAFKxx(1).Number))
    '//2012/12/10 �ی�҂Ɍ_����Ԃ͖����̂Ł��U�֊��Ԃɂ���
    lblCAKYxx(0).Caption = lblCAFKxx(0).Caption
    lblCAKYxx(1).Caption = lblCAFKxx(1).Caption
'//2003/01/31 ���t���O�� NULL �ɂȂ�̂ŕύX
    lblCAKYFG.Caption = Val(chkCAKYFG.Value)
    lblCAUSID.Caption = gdDBS.LoginUserName
    If "" = lblCAADDT.Caption Then
        lblCAADDT.Caption = gdDBS.sysDate
    End If
    lblCAUPDT.Caption = gdDBS.sysDate
'//2007/02/05 UpdateRecord() �ł���ƃG���[���E���Ȃ��̂� Recordset.Update() �ł���悤�ɕύX
'//    Call dbcHogoshaMaster.UpdateRecord
    Call dbcHogoshaMaster.Recordset.Update
'//2004/07/09 �����U�փf�[�^�͋��̂܂܂ɂ��Ă����F�ύX�O�E��̍��ق��Ƃ邽��
#If 0 Then
    '//2003/01/31 �����U�֗\��f�[�^�ւ̍X�V
    sql = "UPDATE tfFurikaeYoteiData SET(" & vbCrLf
    sql = sql & " FAKKBN,FABANK,FASITN,FAKZSB,FAKZNO,FAYBTK,FAYBTN,FAKZNM,FASKGK,FAKYFG,FAUSID,FAUPDT" & vbCrLf
    sql = sql & " ) = (SELECT " & vbCrLf
    sql = sql & " CAKKBN,CABANK,CASITN,CAKZSB,CAKZNO,CAYBTK,CAYBTN,CAKZNM,CASKGK,CAKYFG,CAUSID,CAUPDT" & vbCrLf
    sql = sql & " FROM tcHogoshaMaster" & vbCrLf
    sql = sql & " WHERE CAITKB = FAITKB" & vbCrLf
    sql = sql & "   AND CAKYCD = FAKYCD" & vbCrLf
    sql = sql & "   AND CAHGCD = FAHGCD" & vbCrLf
    sql = sql & "   AND CASQNO =  " & lblCASQNO.Caption & vbCrLf
    sql = sql & " )" & vbCrLf
    sql = sql & " WHERE FAITKB = '" & lblCAITKB.Caption & "'" & vbCrLf
    sql = sql & "   AND FAKYCD = '" & lblCAKYCD.Caption & "'" & vbCrLf
    sql = sql & "   AND FAHGCD = '" & lblCAHGCD.Caption & "'" & vbCrLf
    sql = sql & "   AND FASQNO BETWEEN " & lblCAFKxx(0).Caption & " AND " & lblCAFKxx(1).Caption & vbCrLf
    Call gdDBS.Database.ExecuteSQL(sql)
'//2004/07/09 ���҂̍X�V�ǉ�
    If "0" <> lblCAKYFG.Caption Then
        sql = "UPDATE tfFurikaeYoteiData SET(" & vbCrLf
        sql = sql & " FASKGK,FAKYFG,FAUSID,FAUPDT" & vbCrLf
        sql = sql & " ) = (SELECT " & vbCrLf
        sql = sql & " CASKGK,CAKYFG,CAUSID,CAUPDT" & vbCrLf
        sql = sql & " FROM tcHogoshaMaster" & vbCrLf
        sql = sql & " WHERE CAITKB = FAITKB" & vbCrLf
        sql = sql & "   AND CAKYCD = FAKYCD" & vbCrLf
        sql = sql & "   AND CAHGCD = FAHGCD" & vbCrLf
        sql = sql & "   AND CASQNO =  " & lblCASQNO.Caption & vbCrLf
        sql = sql & " )" & vbCrLf
        sql = sql & " WHERE FAITKB = '" & lblCAITKB.Caption & "'" & vbCrLf
        sql = sql & "   AND FAKYCD = '" & lblCAKYCD.Caption & "'" & vbCrLf
        sql = sql & "   AND FAHGCD = '" & lblCAHGCD.Caption & "'" & vbCrLf
        sql = sql & "   AND FASQNO > " & lblCAFKxx(1).Caption & vbCrLf
        Call gdDBS.Database.ExecuteSQL(sql)
    End If
#End If
    pUpdateRecord = True
    Exit Function
pUpdateRecordError:
    'Call MsgBox("�X�V�������ɃG���[���������܂���." & vbCrLf & vbCrLf & Error, vbCritical + vbOKOnly, mCaption)
    Call gdDBS.ErrorCheck       '//�G���[�g���b�v
End Function

Private Sub cmdUpdate_Click()
    If lblShoriKubun.Caption = eShoriKubun.Delete Then
        Dim sql As String, dyn As OraDynaset
        'Dim sql As String, dyn As Object
        sql = "SELECT FAITKB AS CNT"
        sql = sql & " FROM tfFurikaeYoteiData"
        sql = sql & " WHERE FAITKB = '" & lblCAITKB.Caption & "'"
        sql = sql & "   AND FAKYCD = '" & lblCAKYCD.Caption & "'"
        sql = sql & "   AND FAHGCD = '" & lblCAHGCD.Caption & "'"
        sql = sql & " UNION "
        sql = sql & "SELECT FBITKB AS CNT"
        sql = sql & " FROM tfFurikaeYoteiTran"
        sql = sql & " WHERE FBITKB = '" & lblCAITKB.Caption & "'"
        sql = sql & "   AND FBKYCD = '" & lblCAKYCD.Caption & "'"
        sql = sql & "   AND FBHGCD = '" & lblCAHGCD.Caption & "'"
#If ORA_DEBUG = 1 Then
        Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
        Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
        If Not dyn.EOF() Then
            Call MsgBox("�����U�փf�[�^�Ŏg�p����Ă��邽��" & vbCrLf & vbCrLf & "�폜���鎖�͏o���܂���.", vbCritical, mCaption)
            Exit Sub
        End If
        If vbOK <> MsgBox("�폜���܂����H" & vbCrLf & vbCrLf & "���ɖ߂����Ƃ͏o���܂���.", vbInformation + vbOKCancel + vbDefaultButton2, mCaption) Then
            Exit Sub
        Else
'//2002/11/26 OIP-00000 ORA-04108 �ŃG���[�ɂȂ�̂� Execute() �Ŏ��s����悤�ɕύX.
'// Oracle Data Control 8i(3.6) 9i(4.2) �̈Ⴂ���ȁH
'//            Call dbcHogoshaMaster.Recordset.Delete
            Call dbcHogoshaMaster.UpdateControls
            sql = "DELETE tcHogoshaMaster"
            sql = sql & " WHERE CAITKB = '" & lblCAITKB.Caption & "'"
            sql = sql & "   AND CAKYCD = '" & lblCAKYCD.Caption & "'"
            sql = sql & "   AND CAHGCD = '" & lblCAHGCD.Caption & "'"
            sql = sql & "   AND CASQNO =  " & lblCASQNO.Caption
            Call gdDBS.Database.ExecuteSQL(sql)
        End If
    Else
'//2013/02/26 �����ύX���̍X�V���̒ǉ��X�V�̍ۂɂQ�x pUpdateRecord() �����s�����̂𐧌䂷��
        mRirekiAddNewUpdate = False
        '//���͓��e�`�F�b�N�Ŏ���߂����̂ŏI��
        If False = pUpdateErrorCheck Then
            Exit Sub
        End If
'//2013/02/26 �����ύX���̍X�V���̒ǉ��X�V�̍ۂɂQ�x pUpdateRecord() �����s�����̂𐧌䂷��
        If mRirekiAddNewUpdate = False Then
            If False = pUpdateRecord Then
                Exit Sub
            End If
            '//2015/02/26 �ߋ��̐U�֏I�����ƃ����N������̂œǍ��ݎ��̊J�n����ۊǁA�ύX���ɉߋ��̏I������ύX����
            pRirekiAdjust
        End If
    End If
    Call pLockedControl(True)
    Call cboABKJNM.SetFocus
End Sub

'//2015/02/26 �ߋ��̐U�֏I�����ƃ����N������̂œǍ��ݎ��̊J�n����ۊǁA�ύX���ɉߋ��̏I������ύX����
Private Sub pRirekiAdjust()
    If lblSaveFKST.Caption = lblCAFKxx(0).Caption Then
        Exit Sub
    End If
    Dim sql As String, dyn As OraDynaset
    sql = "SELECT * FROM tcHogoshaMaster"
    sql = sql & " WHERE CAITKB = '" & lblCAITKB.Caption & "'"
    sql = sql & "   AND CAKYCD = '" & lblCAKYCD.Caption & "'"
    sql = sql & "   AND CAHGCD = '" & lblCAHGCD.Caption & "'"
    sql = sql & "   AND CASQNO <  " & lblCASQNO.Caption
    sql = sql & " ORDER BY CASQNO DESC" '//���߂�擪�ɂ���
    Set dyn = gdDBS.OpenRecordset(sql)
    If dyn.RecordCount = 0 Then
        Exit Sub   '//�ߋ��f�[�^���Ȃ��̂ŏI��
    End If
    Dim chgDate As String
    chgDate = Format(DateSerial(txtCAFKxx(0).Year, txtCAFKxx(0).Month, 0), "yyyyMMdd")
    dyn.Edit
    dyn.Fields("CAFKED").Value = chgDate
    dyn.Fields("CAKYED").Value = chgDate
    dyn.Fields("CAUSID").Value = "AdjustCAFKED"
    dyn.Fields("CAUPDT").Value = gdDBS.sysDate()
    dyn.Update
    dyn.Close
End Sub

Private Sub cmdCancel_Click()
    Call dbcHogoshaMaster.UpdateControls
    Call pLockedControl(True)
    Call cboABKJNM.SetFocus
End Sub

Private Sub cmdEnd_Click()
    Call dbcHogoshaMaster.UpdateControls
    Unload Me
End Sub

Private Sub cmdKakutei_Click()
    If dblBankList.Text = "" Or dblShitenList.Text = "" Then
        Exit Sub
    End If
    txtCABANK.Text = Left(dblBankList.Text, 4)
'    lblBankName.Caption = Mid(dblBankList.Text, 6)
    txtCASITN.Text = Left(dblShitenList.Text, 3)
'    lblShitenName.Caption = Mid(dblShitenList.Text, 5)
    cmdKakutei.Enabled = False
End Sub

Private Sub cboBankYomi_Click()
    Call gdDBS.BankDbListRefresh(dbcBank, cboBankYomi, dblBankList, eBankRecordKubun.Bank)
    dbcShiten.RecordSource = ""
    dbcShiten.Refresh
    dblShitenList.ListField = ""
    dblShitenList.Refresh
    cmdKakutei.Enabled = False
End Sub

Private Sub cboShitenYomi_Click()
    If dblBankList.Text = "" Then
        Exit Sub
    End If
    Call gdDBS.BankDbListRefresh(dbcShiten, cboShitenYomi, dblShitenList, eBankRecordKubun.Shiten, Left(dblBankList.Text, 4))
    cmdKakutei.Enabled = False
End Sub

Private Sub dbcHogoshaMaster_Error(DataErr As Integer, Response As Integer)
    Debug.Print
End Sub

Private Sub dblBankList_Click()
    cboShitenYomi.ListIndex = -1
    Call cboShitenYomi_Click
End Sub

Private Sub dblShitenList_Click()
    cmdKakutei.Enabled = dblBankList.Text <> ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    '//��s�ƗX�֋ǂ� Frame �𐮗񂷂�
    fraBank(1).Top = fraBank(0).Top
    fraBank(1).Left = fraBank(0).Left
    fraBank(1).Height = fraBank(0).Height
    fraBank(1).Width = fraBank(0).Width
    'fraBank(0).BackColor = Me.BackColor
    'fraBank(1).BackColor = Me.BackColor
    fraBank(0).BorderStyle = vbBSNone
    fraBank(1).BorderStyle = vbBSNone
    fraBankList.BorderStyle = vbBSNone
    'fraKouzaShubetsu.BackColor = Me.BackColor
    '//�����l���Z�b�g
    optShoriKubun(0).Value = True
 
    dbcBank.RecordSource = ""
    dbcShiten.RecordSource = ""
    dbcHogoshaMaster.RecordSource = ""
    dbcItakushaMaster.RecordSource = "SELECT * FROM taItakushaMaster ORDER BY ABITCD"
    dbcItakushaMaster.ReadOnly = True
    Call pLockedControl(True)
    Call mForm.pInitControl
    '//�_��ҁE�ی�҃R�[�h���͎��͂��̒�`���O��
    'txtCAKYCD.KeyNext = ""
    'txtCAHGCD.KeyNext = ""
    '//�����l���Z�b�g�F�Q�ƃ��[�h
    optShoriKubun(eShoriKubun.Refer).Value = True
    lblBAKJNM.Caption = ""
    spnRireki.Visible = False
    lblBankName.Caption = ""
    lblShitenName.Caption = ""
    Call gdDBS.SetItakushaComboBox(cboABKJNM)
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmHogoshaMaster = Nothing
    Set mForm = Nothing
    If gdForm Is Nothing Then
        End
    Else
        Call gdForm.Show
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub lblCAKKBN_Change()
    optCAKKBN(Val(lblCAKKBN.Caption)).Value = True
End Sub

Private Sub lblCAFKxx_Change(Index As Integer)
    txtCAFKxx(Index).Number = Val(lblCAFKxx(Index).Caption)
End Sub

'Private Sub lblCAKYxx_Change(Index As Integer)
'    txtCAKYxx(Index).Number = Val(lblCAKYxx(Index).Caption)
'End Sub

Private Sub lblCAKZSB_Change()
    optCAKZSB(Val(lblCAKZSB.Caption)).Value = True
End Sub

Private Sub optCAKKBN_Click(Index As Integer)
    fraKinnyuuKikan.Tag = Index
    Call fraBank(Index).ZOrder(0)
    fraBankList.Visible = (Index = 0)
    lblCAKKBN.Caption = Index
    '//�t�H�[�J�X��������̂Őݒ肷��.
    txtCABANK.TabStop = Index = eBankKubun.KinnyuuKikan
    txtCASITN.TabStop = Index = eBankKubun.KinnyuuKikan
    txtCAKZNO.TabStop = Index = eBankKubun.KinnyuuKikan
    txtCAYBTK.TabStop = Index = eBankKubun.YuubinKyoku
    txtCAYBTN.TabStop = Index = eBankKubun.YuubinKyoku
End Sub

Private Sub optCAKZSB_Click(Index As Integer)
    lblCAKZSB.Caption = Index
End Sub

Private Sub optShoriKubun_Click(Index As Integer)
    On Error Resume Next    'Form_Load()���Ƀt�H�[�J�X�𓖂Ă��Ȃ����G���[�ƂȂ�̂ŉ���̃G���[����
    lblShoriKubun.Caption = Index
    Call cboABKJNM.SetFocus
End Sub

Private Sub spnRireki_DownClick()
    '//��̃��R�[�h�Ɉړ�
    If True = gdDBS.MoveRecords(dbcHogoshaMaster, -1) Then '//�f�[�^�� DESC ORDER �������Ă���̂ł���ł悢
        On Error GoTo spnRireki_SpinDownError
'//�ŏI�̃f�[�^�̂ݕҏW�\�Ƃ���
        If optShoriKubun(eShoriKubun.Refer).Value = False Then
            If dbcHogoshaMaster.Recordset.IsFirst Then
                dbcHogoshaMaster.Recordset.Edit     '//�����Ŕr�����|����
                Call pLockedControl(False)
                spnRireki.Visible = True
                '//���̃{�^���͎x�X���N���b�N�������Ɏg����悤�ɂ���.
                cmdKakutei.Enabled = False
            End If
        End If
    Else
        Call MsgBox("����ȍ~�Ƀf�[�^�͂���܂���.", vbInformation, mCaption)
    End If
    Exit Sub
spnRireki_SpinDownError:
    Call gdDBS.ErrorCheck   '//�r������p�G���[�g���b�v
'    Call spnRireki_SpinUp
End Sub

Private Sub spnRireki_UpClick()
    '//�O�̃��R�[�h�Ɉړ�
    If True = gdDBS.MoveRecords(dbcHogoshaMaster, 1) Then '//�f�[�^�� DESC ORDER �������Ă���̂ł���ł悢
'//�ŏI�̃f�[�^�̂ݕҏW�\�Ƃ���
'        dbcKeiyakushaMaster.Recordset.Edit
        If optShoriKubun(eShoriKubun.Refer).Value = False Then
            Call mForm.LockedControlAllTextBox
            cmdEnd.Enabled = True
            cmdCancel.Enabled = True
        End If
    Else
        Call MsgBox("����ȑO�Ƀf�[�^�͂���܂���.", vbInformation, mCaption)
    End If
End Sub

Private Sub txtCABANK_Change()
    If 0 <= Len(Trim(txtCABANK.Text)) And Len(Trim(txtCABANK.Text)) < 4 Then
        lblBankName.Caption = ""
        Exit Sub
    End If
'    Dim dyn As OraDynaset
    Dim dyn As Object
    Set dyn = gdDBS.SelectBankMaster("DISTINCT DAKJNM", eBankRecordKubun.Bank, Trim(txtCABANK.Text), vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblBankName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))
    If "" <> Trim(txtCASITN.Text) Then
        mBankChange = True
        Call txtCASITN_Change   '��Ɏx�X�R�[�h������t���Ďx�X�����\���ł��Ȃ��̂�...�B
    End If
End Sub

Private Sub txtCAFKxx_DropOpen(Index As Integer, NoDefault As Boolean)
    txtCAFKxx(Index).Calendar.Holidays = gdDBS.Holiday(txtCAFKxx(Index).Year)
End Sub

Private Sub txtCAKJNM_Furigana(Yomi As String)
    '//���݂̓ǂ݃J�i���ƌ������`�l���������Ȃ�ǂ݃J�i���ƌ������`�l���ɓ]��
    If Trim(txtCAKNNM.Text) = Trim(txtCAKZNM.Text) Then
        txtCAKNNM.Text = txtCAKNNM.Text & Yomi
        txtCAKZNM.Text = txtCAKNNM.Text
    Else
        txtCAKNNM.Text = txtCAKNNM.Text & Yomi
    End If
End Sub

Public Sub txtCAHGCD_KeyDown(KeyCode As Integer, Shift As Integer)
    '// Return �܂��� Shift�{TAB �̂Ƃ��̂ݏ�������
    If Not (KeyCode = vbKeyReturn) Then
        Exit Sub
    End If
    
'    Dim sql As String, dyn As OraDynaset
    Dim sql As String, dyn As Object
    Dim msg As String
        
    If "" = Trim(txtCAHGCD.Text) Then
        Exit Sub
    End If
    Call txtCAKYCD_KeyDown(KeyCode, Shift)
    '�G���[�̏ꍇ KeyCode = 0 ���Ԃ�
    If KeyCode = 0 Then
        Exit Sub
    End If
'//2006/04/26 �O�[�����ߍ���
    txtCAHGCD.Text = Format(Val(txtCAHGCD.Text), String(8, "0"))
    sql = "SELECT * FROM tcHogoshaMaster"
    sql = sql & " WHERE CAITKB = '" & cboABKJNM.ItemData(cboABKJNM.ListIndex) & "'"
    sql = sql & "   AND CAKYCD = '" & txtCAKYCD.Text & "'"
    sql = sql & "   AND CAHGCD = '" & txtCAHGCD.Text & "'"
    sql = sql & " ORDER BY CASQNO DESC"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If 0 = dyn.RecordCount Then
        If eShoriKubun.Add <> lblShoriKubun.Caption Then     '���R�[�h�����ŐV�K�ȊO�̎�
            msg = "�Y���f�[�^�͑��݂��܂���.(" & lblHogoshaCode.Caption & ")"
        End If
    ElseIf eShoriKubun.Add = lblShoriKubun.Caption Then      '���R�[�h�L��ŐV�K�̎�
        msg = "���Ƀf�[�^�����݂��܂�.(" & lblHogoshaCode.Caption & ")"
    End If
    If msg <> "" Then
        Call MsgBox(msg, vbInformation, mCaption)
        Call txtCAHGCD.SetFocus
        Exit Sub
    End If
    '//��񃁃b�Z�[�W�}�~
    dbcHogoshaMaster.RecordSource = sql
    Call dbcHogoshaMaster.Refresh
    On Error GoTo txtCAHGCD_KeyDownError        '//�r������p�G���[�g���b�v
    If 0& = dbcHogoshaMaster.Recordset.RecordCount Then
        '//�V�K�o�^
        dbcHogoshaMaster.Recordset.AddNew
        lblCAITKB.Caption = cboABKJNM.ItemData(cboABKJNM.ListIndex)
        lblCAKYCD.Caption = txtCAKYCD.Text
        lblCAHGCD.Caption = txtCAHGCD.Text
        lblCASQNO.Caption = gdDBS.sysDate("yyyymmdd")
        lblCAKKBN.Caption = 0
        lblCAKZSB.Caption = 1
        'txtCAKYxx(0).Number = 20000101 '//��U�l��ݒ肵�Ȃ��Ɓu�O�v���Z�b�g����Ȃ��F�s�v�c�H
        'txtCAKYxx(0).Number = gdDBS.sysDate("YYYYMM") & "01"
        'txtCAKYxx(1).Number = gdDBS.LastDay(0)
        txtCAFKxx(0).Number = 20000101 '//��U�l��ݒ肵�Ȃ��Ɓu�O�v���Z�b�g����Ȃ��F�s�v�c�H
        txtCAFKxx(0).Number = gdDBS.sysDate("YYYYMM") & "01"
        txtCAFKxx(1).Number = gdDBS.LastDay(0)
    Else
        '//�C���E�폜
        Call dbcHogoshaMaster.Recordset.MoveFirst
        Call dbcHogoshaMaster.Recordset.Edit
'        Call dbcHogoshaMaster.UpdateRecord
    End If
    '//2015/02/26 �ߋ��̐U�֏I�����ƃ����N������̂œǍ��ݎ��̊J�n����ۊǁA�ύX���ɉߋ��̏I������ύX����
    lblSaveFKST.Caption = Me.lblCAFKxx(0).Caption
    '//�Q�ƂŖ�����΃{�^���̐���J�n
    If False = optShoriKubun(eShoriKubun.Refer).Value Then
        Call pLockedControl(False)
    End If
    spnRireki.Visible = dbcHogoshaMaster.Recordset.RecordCount > 1
    '//���̃{�^���͎x�X���N���b�N�������Ɏg����悤�ɂ���.
    cmdKakutei.Enabled = False
    '//��񃁃b�Z�[�W�}�~
    '//�R���g���[����ی�ҁi�����j�ɂ����������߂ɂ��܂��Ȃ��F���ɕ��@��������Ȃ��H
    Call SendKeys("+{TAB}")
    Exit Sub
txtCAHGCD_KeyDownError:
    Call gdDBS.ErrorCheck(dbcHogoshaMaster.Database)    '//�r������p�G���[�g���b�v
End Sub

Private Sub txtCAFKxx_LostFocus(Index As Integer)
    If Index = 0 Then
        lblCAFKxx(Index).Caption = Format(DateSerial(txtCAFKxx(Index).Year, txtCAFKxx(Index).Month, 1), "yyyyMMdd")
    Else
        lblCAFKxx(Index).Caption = gdDBS.LastDay(Val(gdDBS.Nz(txtCAFKxx(Index).Number)))
    End If
End Sub

'Private Sub txtCAKYxx_DropOpen(Index As Integer, NoDefault As Boolean)
'    txtCAKYxx(Index).Calendar.Holidays = gdDBS.Holiday(txtCAKYxx(Index).Year)
'End Sub

'Private Sub txtCAKYxx_LostFocus(Index As Integer)
'    If txtCAKYxx(Index).Number Then
'        lblCAKYxx(Index).Caption = Format(txtCAKYxx(Index).Year, "0000") & Format(txtCAKYxx(Index).Month, "00") & IIf(Index = 0, "01", "31")
'    End If
'End Sub

Private Sub txtCASITN_Change()
    If 0 <= Len(Trim(txtCASITN.Text)) And Len(Trim(txtCASITN.Text)) < 3 Then
        lblShitenName.Caption = ""
        Exit Sub
    End If
'    Dim dyn As OraDynaset
    Dim dyn As Object
    Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Shiten, Trim(txtCABANK.Text), Trim(txtCASITN.Text), vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"�x�X��_����" �œǂ߂Ȃ�
    '//2006/07/25 �f�[�^�Ȃ��̎��ANULL ������
    If mBankChange And Trim(txtCABANK.Text) <> "" And lblShitenName.Caption = "" Then
        txtCASITN.Text = ""
    End If
    mBankChange = False
End Sub

Private Sub txtCAKYCD_KeyDown(KeyCode As Integer, Shift As Integer)
    '// Return �܂��� Shift�{TAB �̂Ƃ��̂ݏ�������
    If Not (KeyCode = vbKeyReturn) Then
        Exit Sub
    End If
    
'    Dim sql As String, dyn As OraDynaset
    Dim sql As String, dyn As Object
    Dim msg As String
        
    If "" = Trim(txtCAKYCD.Text) Then
        Exit Sub
    End If
'//2006/04/26 �O�[�����ߍ���
    txtCAKYCD.Text = Format(Val(txtCAKYCD.Text), String(7, "0"))
    sql = "SELECT DISTINCT BAITKB,BAKYCD,BAKJNM,BAKYFG FROM tbKeiyakushaMaster"
    sql = sql & " WHERE BAITKB = '" & cboABKJNM.ItemData(cboABKJNM.ListIndex) & "'"
    sql = sql & "   AND BAKYCD = '" & txtCAKYCD.Text & "'"
    sql = sql & "   AND TO_CHAR(SYSDATE,'YYYYMMDD') BETWEEN BAKYST AND BAKYED" '//�L���f�[�^�i����
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If 0 = dyn.RecordCount Then
        Call dyn.Close
        KeyCode = 0
        '//                                        �u�_��Ҕԍ��v
        If dyn.Fields("BAKYFG") <> "0" Then
            Call MsgBox("�Y���f�[�^�����݂��܂���.(" & lblKeiyakushaCode.Caption & ")", vbInformation, mCaption)
        Else
            Call MsgBox("�_��҂�����ԁA�������_����Ԃ��I�����Ă��܂�.(" & lblKeiyakushaCode.Caption & ")", vbInformation, mCaption)
        End If
        Call txtCAKYCD.SetFocus
        Exit Sub
    End If
    lblBAKJNM.Caption = dyn.Fields("BAKJNM")
    Call dyn.Close
End Sub

Private Function pUpdateErrorCheck() As Boolean
'//2006/06/26 �U���݈˗����ɂ��������W�b�N���L��̂Œ���
    '///////////////////////////////
    '//�K�{���͍��ڂƐ������`�F�b�N
    
    Dim obj As Object, msg As String
    '//�ی�ҁE�������͕̂K�{
    If txtCAKJNM.Text = "" Then
        Set obj = txtCAKJNM
        msg = "�������`�l(����)�͕K�{���͂ł�."
    End If
    '//�ی�ҁE�J�i���͕̂K�{
'''    If txtCAKNNM.Text = "" Then
'''        Set obj = txtCAKNNM
'''        msg = "�ی�Җ�(�J�i)�͕K�{���͂ł�."
'''    End If
'    If IsNull(txtCAKYxx(0).Number) Then
'        Set obj = txtCAKYxx(0)
'        msg = "�_����Ԃ̊J�n���͕K�{���͂ł�."
'    ElseIf IsNull(txtCAKYxx(1).Number) Then
'        Set obj = txtCAKYxx(1)
'        msg = "�_����Ԃ̏I�����͕K�{���͂ł�."
'    ElseIf txtCAKYxx(0).Text > txtCAKYxx(1).Text Then
'        Set obj = txtCAKYxx(0)
'        msg = "�_����Ԃ��s���ł�."
    If IsNull(txtCAFKxx(0).Number) Then
        Set obj = txtCAFKxx(0)
        msg = "�U�֊��Ԃ̊J�n���͕K�{���͂ł�."
    ElseIf IsNull(txtCAFKxx(1).Number) Then
        Set obj = txtCAFKxx(1)
        msg = "�U�֊��Ԃ̏I�����͕K�{���͂ł�."
    ElseIf txtCAFKxx(0).Text > txtCAFKxx(1).Text Then
        Set obj = txtCAFKxx(0)
        msg = "�U�֊��Ԃ��s���ł�."
'//2013/06/18 �P�N�ȏ�O�̓��͕s�ɕύX
'    ElseIf txtCAFKxx(1).Number < Val(gdDBS.sysDate("yyyymm")) - 100 & "01" Then
'        Set obj = txtCAFKxx(1)
'        msg = "�U�֊��Ԃ̏I�����������ȑO�ł�."
    ElseIf txtCAFKxx(1).Number < Val(gdDBS.sysDate("yyyymm")) - 100 & "01" Then
        Set obj = txtCAFKxx(1)
        msg = "�U�֊��Ԃ̏I�������P�N�ȏ�O�ł�."
    End If
    
    If lblCAKKBN.Caption = eBankKubun.KinnyuuKikan Then
        If txtCABANK.Text = "" Or lblBankName.Caption = "" Then
            Set obj = txtCABANK
            msg = "���Z�@�ւ͕K�{���͂ł�."
        ElseIf txtCASITN.Text = "" Or lblShitenName.Caption = "" Then
            Set obj = txtCASITN
            msg = "�x�X�͕K�{���͂ł�."
        ElseIf Not (lblCAKZSB.Caption = eBankYokinShubetsu.Futsuu _
                 Or lblCAKZSB.Caption = eBankYokinShubetsu.Touza) Then
            Set obj = optCAKZSB(eBankYokinShubetsu.Futsuu)
            msg = "�a����ʂ͕K�{���͂ł�."
        ElseIf Len(Trim(txtCAKZNO.Text)) <> txtCAKZNO.MaxLength Then
            Set obj = txtCAKZNO
            msg = "�����ԍ��͂V���K�{���͂ł�."
        End If
    ElseIf lblCAKKBN.Caption = eBankKubun.YuubinKyoku Then
        If Len(Trim(txtCAYBTK.Text)) <> txtCAYBTK.MaxLength Then
            Set obj = txtCAYBTK
            msg = "�ʒ��L���͂R���K�{���͂ł�."
        ElseIf Len(Trim(txtCAYBTN.Text)) <> txtCAYBTN.MaxLength Then
            Set obj = txtCAYBTN
            msg = "�ʒ��ԍ��͂W���K�{���͂ł�."
        ElseIf "1" <> Right(txtCAYBTN.Text, 1) Then
'//2006/04/26 �����ԍ��`�F�b�N
            Set obj = txtCAYBTN
            msg = "�ʒ��ԍ��̖������u�P�v�ȊO�ł�."
        End If
    End If
    If Trim(txtCAKZNM.Text) = "" Then
        Set obj = txtCAKZNM
        msg = "�������`�l(�J�i)�͕K�{���͂ł�."
    End If
    '//Object ���ݒ肳��Ă��邩�H
    If TypeName(obj) <> "Nothing" Then
        Call MsgBox(msg, vbCritical, mCaption)
        Call obj.SetFocus
        Exit Function
    End If
    
    If lblCASQNO.Caption = gdDBS.sysDate("yyyymmdd") Then
        pUpdateErrorCheck = True    '//�r�d�p���{���Ȃ̂ł��̂܂܍X�V
        Exit Function
    End If
    pUpdateErrorCheck = pRirekiAddNew()
    Exit Function
pUpdateErrorCheckError:
    Call gdDBS.ErrorCheck       '//�G���[�g���b�v
    pUpdateErrorCheck = False   '//���S�̂��߁FFalse �ŏI������͂�
End Function

Private Function pRirekiAddNew()
'    Dim sql As String, dyn As OraDynaset
    Dim sql As String, dyn As Object
    Dim AddRireki As String
    
    sql = "SELECT * FROM tcHogoshaMaster"
    sql = sql & " WHERE CAITKB = '" & lblCAITKB.Caption & "'"
    sql = sql & "   AND CAKYCD = '" & lblCAKYCD.Caption & "'"
    sql = sql & "   AND CAHGCD = '" & lblCAHGCD.Caption & "'"
    sql = sql & "   AND CASQNO =  " & lblCASQNO.Caption
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If dyn.EOF Then
        Exit Function   '//�V�K�o�^�Ȃ̂Ń`�F�b�N����
    End If
        
    If txtCAKJNM.Text <> gdDBS.Nz(dyn.Fields("CAKJNM")) _
    Or txtCAKZNM.Text <> gdDBS.Nz(dyn.Fields("CAKZNM")) Then
'''    If txtCAKJNM.Text <> gdDBS.Nz(dyn.Fields("CAKJNM")) _
'''    Or txtCAKNNM.Text <> gdDBS.Nz(dyn.Fields("CAKNNM")) Then
        AddRireki = "�������`�l"
    ElseIf lblCAKKBN.Caption <> gdDBS.Nz(dyn.Fields("CAKKBN")) Then
        AddRireki = "�U�֌���"
    ElseIf lblCAKKBN.Caption = eBankKubun.KinnyuuKikan Then
        '//���Z�@�֏�񂪈Ⴆ�Η������ǉ�
        If txtCABANK.Text <> gdDBS.Nz(dyn.Fields("CABANK")) _
         Or txtCASITN.Text <> gdDBS.Nz(dyn.Fields("CASITN")) _
         Or lblCAKZSB.Caption <> gdDBS.Nz(dyn.Fields("CAKZSB")) _
         Or txtCAKZNO.Text <> gdDBS.Nz(dyn.Fields("CAKZNO")) Then
            AddRireki = "���ԋ��Z�@��"
        End If
    ElseIf lblCAKKBN.Caption = eBankKubun.YuubinKyoku Then
        '//�X�֋Ǐ�񂪈Ⴆ�Η������ǉ�
        If txtCAYBTK.Text <> gdDBS.Nz(dyn.Fields("CAYBTK")) _
         Or txtCAYBTN.Text <> gdDBS.Nz(dyn.Fields("CAYBTN")) Then
            AddRireki = "�X�֋�"
        End If
'''    ElseIf txtCAKZNM.Text <> gdDBS.Nz(dyn.Fields("CAKZNM")) Then
'''        AddRireki = "�������`�l"
    End If
    
    '///////////////////////////
    '//�����쐬���Ȃ��ꍇ�I��
    If "" = AddRireki Then
        pRirekiAddNew = True    '//���݂̃��R�[�h�ɍX�V
        Exit Function
    End If
    
    '///////////////////////////////////////////
    '//�ύX���e��`�̉�ʂ�\������
    Load frmMakeNewData
    With frmMakeNewData
        '//�t�H�[�������̃t�H�[���̒����Ɉʒu�t������
        .Top = Me.Top + (Me.Height - .Height) / 2
        .Left = Me.Left + (Me.Width - .Width) / 2
        .lblMessage.Caption = "�u" & AddRireki & "�v�̏�񂪕ύX���ꂽ���ߗ������쐬���܂�." & vbCrLf & vbCrLf & _
                              "�u�ǉ��v�@�����Ƃ��ĉߋ��̏����c���ꍇ�͂��̃{�^���������܂�." & vbCrLf & _
                              "�@�@�@�@�@��Ɂu�����ύX�v�̍ۂɂ��̑�������ĉ�����." & vbCrLf & vbCrLf & _
                              "�u�㏑���v���݂̃f�[�^�ɏ㏑������ꍇ�͂��̃{�^���������܂�." & vbCrLf & _
                              "�@�@�@�@�@��Ɂu��������v�̍ۂɂ��̑�������ĉ�����."
        .lblFurikomi.Caption = "�U�֊J�n��"
        Call .Show(vbModal)
        '//���j������邩�킩��Ȃ��̂Ń��[�J���R�s�[���Ă���
        Dim PushButton As Integer, KeiyakuEnd As Long, FurikaeEnd As Long
        PushButton = .mPushButton
        '//2012/12/10 �_����ԁ��U�֊��ԂƂ���
        KeiyakuEnd = .mFurikaeEnd   '.mKeiyakuEnd
        FurikaeEnd = .mFurikaeEnd
        Set frmMakeNewData = Nothing
        If PushButton = ePushButton.Update Then
            pRirekiAddNew = True    '//���݂̃��R�[�h�ɍX�V�F���̎������߂��čX�V����.
            Exit Function
        ElseIf PushButton = ePushButton.Cancel Then
            Exit Function
        End If
    End With
    '//���������ʓ��e�̍X�V�y�ї����쐬�J�n
    
    '//�O�����Ēǉ����郌�R�[�h�폜
    sql = " DELETE tcHogoshaMaster"
    sql = sql & " WHERE CAITKB = '" & lblCAITKB.Caption & "'"
    sql = sql & "   AND CAKYCD = '" & lblCAKYCD.Caption & "'"
    sql = sql & "   AND CAHGCD = '" & lblCAHGCD.Caption & "'"
    sql = sql & "   AND CASQNO = -1"
    Call gdDBS.Database.ExecuteSQL(sql)
    
    '////////////////////////////////////////////////
    '//�e�[�u����`���ύX���ꂽ�ꍇ���ӂ��邱�ƁI�I
    Dim FixedCol As String
    FixedCol = "CAITKB,CAKYCD,CAKJNM,CAHGCD,CAKNNM," & _
               "CASTNM,CAKKBN,CABANK,CASITN,CAKZSB,CAKZNO," & _
               "CAKZNM,CAYBTK,CAYBTN,CAKYST,CAFKST,CASKGK," & _
               "CAHKGK,CAKYDT,CAKYFG,CATRFG,CAADDT,CAUSID," & _
               "canwdt"
    '���݂̍X�V�O�f�[�^�ޔ�
    sql = "INSERT INTO tcHogoshaMaster("
    sql = sql & "CASQNO,CAKYED,CAFKED,CAUPDT,"
    sql = sql & FixedCol
    sql = sql & ") SELECT "
    sql = sql & "-1,"
    '//���͂��ꂽ���̑O��������ݒ�
    sql = sql & "TO_CHAR(TO_DATE(" & KeiyakuEnd & ",'YYYYMMDD')-1,'YYYYMMDD'),"
    sql = sql & "TO_CHAR(TO_DATE(" & FurikaeEnd & ",'YYYYMMDD')-1,'YYYYMMDD'),"
    sql = sql & " SYSDATE,"
    sql = sql & FixedCol
    sql = sql & " FROM tcHogoshaMaster"
    sql = sql & " WHERE CAITKB = '" & lblCAITKB.Caption & "'"
    sql = sql & "   AND CAKYCD = '" & lblCAKYCD.Caption & "'"
    sql = sql & "   AND CAHGCD = '" & lblCAHGCD.Caption & "'"
    sql = sql & "   AND CASQNO =  " & lblCASQNO.Caption
    Call gdDBS.Database.ExecuteSQL(sql)
    
    'txtCAKYxx(0).Number = KeiyakuEnd
    txtCAFKxx(0).Number = FurikaeEnd
    
    '//��ʂ̓��e���X�V:cmdUpdate()�̈ꕔ�֐������s
    Call pUpdateRecord
    
    On Error GoTo pRirekiAddNewError
    '//��ʂ̃f�[�^�̂r�d�p��{���ɂ���
    sql = "UPDATE tcHogoshaMaster SET "
    sql = sql & "CASQNO = TO_CHAR(SYSDATE,'YYYYMMDD'),"
    sql = sql & "CAUSID = '" & gdDBS.LoginUserName & "',"
    sql = sql & "CAUPDT = SYSDATE"
    sql = sql & " WHERE CAITKB = '" & lblCAITKB.Caption & "'"
    sql = sql & "   AND CAKYCD = '" & lblCAKYCD.Caption & "'"
    sql = sql & "   AND CAHGCD = '" & lblCAHGCD.Caption & "'"
    sql = sql & "   AND CASQNO =  " & lblCASQNO.Caption
    Call gdDBS.Database.ExecuteSQL(sql)
    '//�ޔ������f�[�^�̂r�d�p��ύX�O�ɂ���
    sql = "UPDATE tcHogoshaMaster SET "
    sql = sql & "CASQNO = " & lblCASQNO.Caption & ","
    sql = sql & "CAUSID = '" & gdDBS.LoginUserName & "',"
    sql = sql & "CAUPDT = SYSDATE"
    sql = sql & " WHERE CAITKB = '" & lblCAITKB.Caption & "'"
    sql = sql & "   AND CAKYCD = '" & lblCAKYCD.Caption & "'"
    sql = sql & "   AND CAHGCD = '" & lblCAHGCD.Caption & "'"
    sql = sql & "   AND CASQNO = -1"
    Call gdDBS.Database.ExecuteSQL(sql)
'//2013/02/26 �����ύX���̍X�V���̒ǉ��X�V�̍ۂɂQ�x pUpdateRecord() �����s�����̂𐧌䂷��
    mRirekiAddNewUpdate = True
    pRirekiAddNew = True
    Exit Function
pRirekiAddNewError:
    Call gdDBS.ErrorCheck       '//�G���[�g���b�v
    pRirekiAddNew = False   '//���S�̂��߁FFalse �ŏI������͂�
End Function

Private Sub mnuEnd_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

Private Sub txtCASITN_KeyDown(KeyCode As Integer, Shift As Integer)
    mBankChange = True
End Sub

Private Sub txtCAYBTK_LostFocus()
'//2006/04/26 �O�[�����ߍ���
    If "" <> txtCAYBTK.Text Then
        txtCAYBTK.Text = Format(Val(txtCAYBTK.Text), "000")
    End If
End Sub

Private Sub txtCAYBTN_LostFocus()
    '//2006/04/26 �O�[�����ߍ���
    If "" <> txtCAYBTN.Text Then
        If "1" <> Right(txtCAYBTN.Text, 1) Then
            Call MsgBox("�������u�P�v�ȊO�ł�.(" & lblTsuchoBango.Caption & ")", vbCritical, mCaption)
        Else
            txtCAYBTN.Text = Format(Val(txtCAYBTN.Text), "00000000")
        End If
    End If
End Sub
