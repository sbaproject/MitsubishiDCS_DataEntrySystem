VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHogoshaMaster 
   Caption         =   "�ی�҃}�X�^�����e�i���X"
   ClientHeight    =   7335
   ClientLeft      =   1710
   ClientTop       =   4725
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
   ScaleHeight     =   7335
   ScaleWidth      =   10125
   Begin VB.CommandButton cmdClassNoChange 
      Caption         =   "�����ԍ�(&Z)"
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
      Left            =   4080
      TabIndex        =   86
      Top             =   6720
      Width           =   1395
   End
   Begin MSComCtl2.UpDown spnRireki 
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      ToolTipText     =   "�O��̗����Ɉړ�"
      Top             =   1860
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   873
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
      TabIndex        =   4
      TabStop         =   0   'False
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
      TabIndex        =   21
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
         TabIndex        =   31
         Top             =   1020
         Width           =   4035
         Begin imText6Ctl.imText txtCAYBTK 
            DataField       =   "CAYBTK"
            DataSource      =   "dbcHogoshaMaster"
            Height          =   285
            Left            =   1860
            TabIndex        =   32
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
         Begin imText6Ctl.imText txtCAYBTN 
            DataField       =   "CAYBTN"
            DataSource      =   "dbcHogoshaMaster"
            Height          =   285
            Left            =   1860
            TabIndex        =   33
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
            LengthAsByte    =   -1
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
            TabIndex        =   57
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
            TabIndex        =   56
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
         TabIndex        =   24
         Top             =   420
         Width           =   3855
         Begin imText6Ctl.imText txtCAKZNO 
            DataField       =   "CAKZNO"
            DataSource      =   "dbcHogoshaMaster"
            Height          =   285
            Left            =   1140
            TabIndex        =   30
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
            LengthAsByte    =   -1
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
            TabIndex        =   26
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
         Begin imText6Ctl.imText txtCABANK 
            DataField       =   "CABANK"
            DataSource      =   "dbcHogoshaMaster"
            Height          =   285
            Left            =   1200
            TabIndex        =   25
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
            TabIndex        =   58
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
               TabIndex        =   28
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
               TabIndex        =   27
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
               TabIndex        =   29
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
               TabIndex        =   59
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
            TabIndex        =   65
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
            TabIndex        =   64
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
            TabIndex        =   63
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
            TabIndex        =   62
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
            TabIndex        =   61
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
            TabIndex        =   60
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
         TabIndex        =   23
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   1395
      End
      Begin imText6Ctl.imText txtCAKZNM 
         DataField       =   "CAKZNM"
         DataSource      =   "dbcHogoshaMaster"
         Height          =   285
         Left            =   420
         TabIndex        =   34
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
         Enabled         =   0
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
         LengthAsByte    =   -1
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
      Begin VB.Label lblKouzaName 
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
         TabIndex        =   83
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
         TabIndex        =   66
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkCAKYFG 
      Caption         =   "���"
      DataField       =   "CAKYFG"
      Height          =   315
      Left            =   4200
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3660
      Width           =   675
   End
   Begin VB.ComboBox cboCAKSCDz 
      BackColor       =   &H000000FF&
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
      Left            =   2880
      List            =   "�ی�҃}�X�^�����e�i���X.frx":05B4
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   85
      Tag             =   "InputKey"
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin imNumber6Ctl.imNumber txtCASKGK 
      DataField       =   "CASKGK"
      DataSource      =   "dbcHogoshaMaster"
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Top             =   4620
      Width           =   1095
      _Version        =   65537
      _ExtentX        =   1931
      _ExtentY        =   503
      Calculator      =   "�ی�҃}�X�^�����e�i���X.frx":05D2
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":05F2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":065E
      Keys            =   "�ی�҃}�X�^�����e�i���X.frx":067C
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":06C6
      Spin            =   "�ی�҃}�X�^�����e�i���X.frx":06E2
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0; -##,###,##0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0; -##,###,##0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999999
      MinValue        =   -99999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   1234567
      MaxValueVT      =   5
      MinValueVT      =   5
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
      TabIndex        =   35
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
         ItemData        =   "�ی�҃}�X�^�����e�i���X.frx":070A
         Left            =   1500
         List            =   "�ی�҃}�X�^�����e�i���X.frx":072F
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   36
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
         ItemData        =   "�ی�҃}�X�^�����e�i���X.frx":0771
         Left            =   3900
         List            =   "�ی�҃}�X�^�����e�i���X.frx":0796
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   38
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
            Size            =   9
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
            Size            =   9
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
         Bindings        =   "�ی�҃}�X�^�����e�i���X.frx":07D8
         Height          =   2040
         Left            =   120
         TabIndex        =   37
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
         Bindings        =   "�ی�҃}�X�^�����e�i���X.frx":07EE
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
         TabIndex        =   68
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
         TabIndex        =   67
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
      TabIndex        =   42
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
      TabIndex        =   0
      Tag             =   "InputKey"
      Top             =   120
      Width           =   3675
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
         Left            =   2820
         TabIndex        =   87
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
         Left            =   1080
         TabIndex        =   2
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
         Left            =   1980
         TabIndex        =   3
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
         Left            =   180
         TabIndex        =   1
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
         Left            =   1500
         TabIndex        =   54
         Top             =   0
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
      TabIndex        =   41
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
      TabIndex        =   43
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
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   "kumon"
      Connect         =   "kumon/kumon"
      RecordSource    =   "SELECT * FROM tcHogoshaMaster"
   End
   Begin imDate6Ctl.imDate txtCAKYxx 
      DataField       =   "CAKYST"
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   14
      Top             =   3660
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   556
      Calendar        =   "�ی�҃}�X�^�����e�i���X.frx":0806
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":0986
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":09F2
      Keys            =   "�ی�҃}�X�^�����e�i���X.frx":0A10
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":0A6E
      Spin            =   "�ی�҃}�X�^�����e�i���X.frx":0A8A
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
   Begin imDate6Ctl.imDate txtCAKYxx 
      DataField       =   "CAKYED"
      Height          =   315
      Index           =   1
      Left            =   3120
      TabIndex        =   15
      Top             =   3660
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   556
      Calendar        =   "�ی�҃}�X�^�����e�i���X.frx":0AB2
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":0C32
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":0C9E
      Keys            =   "�ی�҃}�X�^�����e�i���X.frx":0CBC
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":0D1A
      Spin            =   "�ی�҃}�X�^�����e�i���X.frx":0D36
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
   Begin imDate6Ctl.imDate txtCAFKxx 
      DataField       =   "CAFKST"
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   17
      Top             =   4080
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   556
      Calendar        =   "�ی�҃}�X�^�����e�i���X.frx":0D5E
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":0EDE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":0F4A
      Keys            =   "�ی�҃}�X�^�����e�i���X.frx":0F68
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":0FC6
      Spin            =   "�ی�҃}�X�^�����e�i���X.frx":0FE2
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
   Begin imText6Ctl.imText txtCAKJNM 
      DataField       =   "CAKJNM"
      DataSource      =   "dbcHogoshaMaster"
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Top             =   2460
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":100A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":1076
      Key             =   "�ی�҃}�X�^�����e�i���X.frx":1094
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":10D8
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
      TabIndex        =   5
      Tag             =   "InputKey"
      Top             =   1320
      Width           =   615
      _Version        =   65537
      _ExtentX        =   1085
      _ExtentY        =   503
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":10F4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":1160
      Key             =   "�ی�҃}�X�^�����e�i���X.frx":117E
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":11C2
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
      MaxLength       =   5
      LengthAsByte    =   -1
      Text            =   "12345"
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
      TabIndex        =   7
      Tag             =   "InputKey"
      Top             =   2040
      Width           =   495
      _Version        =   65537
      _ExtentX        =   873
      _ExtentY        =   503
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":11DE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":124A
      Key             =   "�ی�҃}�X�^�����e�i���X.frx":1268
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":12AC
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
         Size            =   9
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
   Begin imNumber6Ctl.imNumber txtCAHKGK 
      DataField       =   "CAHKGK"
      DataSource      =   "dbcHogoshaMaster"
      Height          =   285
      Left            =   1800
      TabIndex        =   20
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65537
      _ExtentX        =   1931
      _ExtentY        =   503
      Calculator      =   "�ی�҃}�X�^�����e�i���X.frx":12C8
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":12E8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":1354
      Keys            =   "�ی�҃}�X�^�����e�i���X.frx":1372
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":13BC
      Spin            =   "�ی�҃}�X�^�����e�i���X.frx":13D8
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0; -##,###,##0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0; -##,###,##0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999999
      MinValue        =   -99999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   1234567
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin imText6Ctl.imText txtCASTNM 
      DataField       =   "CASTNM"
      DataSource      =   "dbcHogoshaMaster"
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   3180
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":1400
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":146C
      Key             =   "�ی�҃}�X�^�����e�i���X.frx":148A
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":14CE
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
      TabIndex        =   12
      Top             =   2820
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":14EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":1556
      Key             =   "�ی�҃}�X�^�����e�i���X.frx":1574
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":15B8
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
      LengthAsByte    =   -1
      Text            =   "�żҲ..................................*"
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
      TabIndex        =   18
      Top             =   4080
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   556
      Calendar        =   "�ی�҃}�X�^�����e�i���X.frx":15D4
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":1754
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":17C0
      Keys            =   "�ی�҃}�X�^�����e�i���X.frx":17DE
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":183C
      Spin            =   "�ی�҃}�X�^�����e�i���X.frx":1858
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
   Begin imText6Ctl.imText txtCAKSCD 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Tag             =   "InputKey"
      Top             =   1680
      Width           =   375
      _Version        =   65537
      _ExtentX        =   661
      _ExtentY        =   503
      Caption         =   "�ی�҃}�X�^�����e�i���X.frx":1880
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����e�i���X.frx":18EC
      Key             =   "�ی�҃}�X�^�����e�i���X.frx":190A
      MouseIcon       =   "�ی�҃}�X�^�����e�i���X.frx":194E
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
      TabIndex        =   84
      Top             =   2820
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
      TabIndex        =   82
      Top             =   3225
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�ύX����z"
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
      TabIndex        =   81
      Top             =   5040
      Visible         =   0   'False
      Width           =   1275
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
      TabIndex        =   80
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
      Left            =   4440
      TabIndex        =   79
      Top             =   4080
      Width           =   375
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
      Left            =   3000
      TabIndex        =   78
      Top             =   5940
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
      Left            =   1980
      TabIndex        =   77
      Top             =   5940
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
      Left            =   3000
      TabIndex        =   76
      Top             =   5640
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
      Left            =   1980
      TabIndex        =   75
      Top             =   5640
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
      TabIndex        =   74
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
      TabIndex        =   73
      Top             =   7200
      Width           =   1755
   End
   Begin VB.Label Label6 
      Alignment       =   1  '�E����
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
      Left            =   360
      TabIndex        =   72
      Tag             =   "InputKey"
      Top             =   1710
      Width           =   1275
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
      Left            =   4020
      TabIndex        =   71
      Top             =   540
      Width           =   975
   End
   Begin VB.Label lblCAKSCD 
      BackColor       =   &H000000FF&
      Caption         =   "�����ԍ�"
      DataField       =   "CAKSCD"
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
      Left            =   3600
      TabIndex        =   70
      Top             =   1020
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
      Height          =   255
      Left            =   360
      TabIndex        =   69
      Tag             =   "InputKey"
      Top             =   900
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
      Left            =   3600
      TabIndex        =   10
      Top             =   2100
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
      Left            =   4020
      TabIndex        =   55
      Top             =   780
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
      Left            =   3600
      TabIndex        =   9
      Top             =   1800
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
      TabIndex        =   53
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label17 
      Alignment       =   1  '�E����
      Caption         =   "�����\��z"
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
      Top             =   4620
      Width           =   1275
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
      Left            =   2520
      TabIndex        =   51
      Tag             =   "InputKey"
      Top             =   1380
      Width           =   2355
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
      Left            =   2820
      TabIndex        =   50
      Top             =   4140
      Width           =   255
   End
   Begin VB.Label Label7 
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
      Left            =   2820
      TabIndex        =   49
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label lblKeiyakushaCode 
      Alignment       =   1  '�E����
      Caption         =   "�_��Ҕԍ�"
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
      TabIndex        =   48
      Tag             =   "InputKey"
      Top             =   1320
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
      Height          =   255
      Left            =   360
      TabIndex        =   47
      Tag             =   "InputKey"
      Top             =   2040
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
      Left            =   290
      TabIndex        =   46
      Top             =   2505
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
      TabIndex        =   45
      Top             =   4140
      Width           =   1275
   End
   Begin VB.Label Label16 
      Alignment       =   1  '�E����
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
      Left            =   360
      TabIndex        =   44
      Top             =   3720
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
Private mIsActivated As Boolean
'//2013/02/26 �����ύX���̍X�V���̒ǉ��X�V�̍ۂɂQ�x pUpdateRecord() �����s�����̂𐧌䂷��
Private mRirekiAddNewUpdate As Boolean

'//2007/06/07 �X�V�E���~�{�^�������S�P�ƂɃR���g���[��
Private Sub pButtonControl(ByVal vMode As Boolean, Optional vExec As Boolean = False)
    If True = mIsActivated Or True = vExec Then
        cmdUpdate.Visible = vMode
        cmdCancel.Visible = vMode
        cmdUpdate.Enabled = vMode
        cmdCancel.Enabled = vMode
        cmdEnd.Enabled = Not vMode
        mIsActivated = True
    End If
    '//�C�����ȊO�͋����ԍ��{�^���̉����͕s�\�ɂ���
    If optShoriKubun(eShoriKubun.Edit).Value And cmdClassNoChange.Enabled Then
        cmdClassNoChange.Visible = Not cmdUpdate.Visible
    Else
        cmdClassNoChange.Visible = False
    End If
End Sub

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
    'cmdEnd.Enabled = blMode
    spnRireki.Visible = False
    '//2007/06/07 �������`�l�͏�ɓ��͂��Ȃ��F�ی�Җ�(�J�i)���R�s�[����l�Ɏd�l�ύX
    txtCAKZNM.Enabled = False
    lblKouzaName.Enabled = False
    cmdKakutei.Enabled = Not blMode
End Sub

#If 0 Then
Private Sub cboCAKSCDz_GotFocus()
    '//TAB �L�[���͎� txtCAKYCD_KeyDown() �C�x���g���������Ȃ��̂�
    Call cboCAKYCDz_KeyDown(vbKeyReturn, 0)
End Sub
#End If

Private Sub chkCAKYFG_Click()
    lblCAKYFG.Caption = chkCAKYFG.Value
    Call pButtonControl(True)
End Sub

Private Sub chkCAKYFG_KeyDown(KeyCode As Integer, Shift As Integer)
    '//���t���O��ݒ肵���̂ŏI�����̓��͂𑣂�.
    '//KeyCode & Shift ���N���A���Ȃ��ƃo�b�t�@�Ɏc��H
    KeyCode = 0
    Shift = 0
    chkCAKYFG.Value = Choose(chkCAKYFG.Value + 1, 1, 0, 0)  '// Index=1,2,3
    Call MsgBox("���̕ύX�����m���܂����B" & vbCrLf & vbCrLf & "�_����ԋy�ѐU�֊��� �I�����̍Đݒ�����ĉ�����.", vbInformation + vbOKOnly, mCaption)
    Call txtCAKYxx(1).SetFocus
End Sub

Private Sub chkCAKYFG_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '//���t���O��ݒ肵���̂ŏI�����̓��͂𑣂�.
    If Button = vbLeftButton Then
        Call chkCAKYFG_KeyDown(vbKeySpace, 0)
    End If
End Sub

Private Sub cmdClassNoChange_Click()
    Load frmClassNoChange
    With frmClassNoChange
        '//�\���t�H�[�������̃t�H�[���̒����ɂ���.
        .Top = Me.Top + (Me.Height - .Height) / 2
        .Left = Me.Left + (Me.Width - .Width) / 2
        .lblCAITKB.Caption = lblCAITKB.Caption
        .lblCAKYCD.Caption = lblCAKYCD.Caption
        .lblCAKSCD.Caption = lblCAKSCD.Caption
        .lblCAHGCD.Caption = lblCAHGCD.Caption
        .txtCAKSCD.Text = ""
        Call .Show(vbModal)
        '//�ύX���ꂽ�̂� .mNewCode �ɒl������
        If "" <> Trim(.mNewCode) Then
            txtCAKSCD.Text = .mNewCode
            'lblCAKSCD.Caption = .mNewCode
            '//�r����������̂� frmClassNoChange() �ōX�V�O�ɕی�҃}�X�^�E���b�N�������Ă���F�X�V�O�Ȃ�s�v
            Call cmdCancel_Click
            'Call pButtonControl(False)
        End If
    End With
End Sub

Private Sub Form_Activate()
    If False = mIsActivated Then
        Call pButtonControl(False, True)
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
    'sql = sql & "   AND CAKSCD = '" & lblCAKSCD.Caption & "'"
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
    lblCAKYxx(0).Caption = Val(gdDBS.Nz(txtCAKYxx(0).Number))
    lblCAKYxx(1).Caption = Val(gdDBS.Nz(txtCAKYxx(1).Number))
    lblCAFKxx(0).Caption = Val(gdDBS.Nz(txtCAFKxx(0).Number))
    lblCAFKxx(1).Caption = Val(gdDBS.Nz(txtCAFKxx(1).Number))
'//2003/01/31 ���t���O�� NULL �ɂȂ�̂ŕύX
    lblCAKYFG.Caption = Val(chkCAKYFG.Value)
    lblCAUSID.Caption = gdDBS.LoginUserName
    If "" = lblCAADDT.Caption Then
        lblCAADDT.Caption = gdDBS.sysDate
    End If
    lblCAUPDT.Caption = gdDBS.sysDate
    Call dbcHogoshaMaster.UpdateRecord
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
    sql = sql & "   AND CAKSCD = FAKSCD" & vbCrLf
    sql = sql & "   AND CAHGCD = FAHGCD" & vbCrLf
    sql = sql & "   AND CASQNO =  " & lblCASQNO.Caption & vbCrLf
    sql = sql & " )" & vbCrLf
    sql = sql & " WHERE FAITKB = '" & lblCAITKB.Caption & "'" & vbCrLf
    sql = sql & "   AND FAKYCD = '" & lblCAKYCD.Caption & "'" & vbCrLf
    sql = sql & "   AND FAKSCD = '" & lblCAKSCD.Caption & "'" & vbCrLf
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
        sql = sql & "   AND CAKSCD = FAKSCD" & vbCrLf
        sql = sql & "   AND CAHGCD = FAHGCD" & vbCrLf
        sql = sql & "   AND CASQNO =  " & lblCASQNO.Caption & vbCrLf
        sql = sql & " )" & vbCrLf
        sql = sql & " WHERE FAITKB = '" & lblCAITKB.Caption & "'" & vbCrLf
        sql = sql & "   AND FAKYCD = '" & lblCAKYCD.Caption & "'" & vbCrLf
        sql = sql & "   AND FAKSCD = '" & lblCAKSCD.Caption & "'" & vbCrLf
        sql = sql & "   AND FAHGCD = '" & lblCAHGCD.Caption & "'" & vbCrLf
        sql = sql & "   AND FASQNO > " & lblCAFKxx(1).Caption & vbCrLf
        Call gdDBS.Database.ExecuteSQL(sql)
    End If
#End If
    pUpdateRecord = True
End Function

Private Sub cmdUpdate_Click()
    If lblShoriKubun.Caption = eShoriKubun.Delete Then
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
        sql = "SELECT FAITKB AS CNT"
        sql = sql & " FROM tfFurikaeYoteiData"
        sql = sql & " WHERE FAITKB = '" & lblCAITKB.Caption & "'"
        sql = sql & "   AND FAKYCD = '" & lblCAKYCD.Caption & "'"
        sql = sql & "   AND FAKSCD = '" & lblCAKSCD.Caption & "'"
        sql = sql & "   AND FAHGCD = '" & lblCAHGCD.Caption & "'"
        sql = sql & " UNION "
        sql = sql & "SELECT FBITKB AS CNT"
        sql = sql & " FROM tfFurikaeYoteiTran"
        sql = sql & " WHERE FBITKB = '" & lblCAITKB.Caption & "'"
        sql = sql & "   AND FBKYCD = '" & lblCAKYCD.Caption & "'"
        sql = sql & "   AND FBKSCD = '" & lblCAKSCD.Caption & "'"
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
            sql = sql & "   AND CAKSCD = '" & lblCAKSCD.Caption & "'"
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
        If False = pUpdateRecord Then
            Exit Sub
        End If
    End If
    Call pLockedControl(True)
    Call txtCAKYCD.SetFocus ' cboABKJNM.SetFocus
    Call pButtonControl(False)
    cmdClassNoChange.Visible = False    '//�����ԍ��C���s�I
End Sub

Private Sub cmdCancel_Click()
    Call dbcHogoshaMaster.UpdateControls
    Call pLockedControl(True)
    Call txtCAKYCD.SetFocus ' cboABKJNM.SetFocus
    Call pButtonControl(False)
    cmdClassNoChange.Visible = False    '//�����ԍ��C���s�I
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
    lblBankName.Caption = Mid(dblBankList.Text, 6)
    txtCASITN.Text = Left(dblShitenList.Text, 3)
    lblShitenName.Caption = Mid(dblShitenList.Text, 5)
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
    '//�����l���Z�b�g�F�C�����[�h
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

Private Sub lblCAKYxx_Change(Index As Integer)
    txtCAKYxx(Index).Number = Val(lblCAKYxx(Index).Caption)
End Sub

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
    Call pButtonControl(True)
End Sub

Private Sub optCAKZSB_Click(Index As Integer)
    lblCAKZSB.Caption = Index
    Call pButtonControl(True)
End Sub

Private Sub optShoriKubun_Click(Index As Integer)
    On Error Resume Next    'Form_Load()���Ƀt�H�[�J�X�𓖂Ă��Ȃ����G���[�ƂȂ�̂ŉ���̃G���[����
    lblShoriKubun.Caption = Index
    Call txtCAKYCD.SetFocus ' cboABKJNM.SetFocus
End Sub

Private Sub SetBankAndShiten()
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Set dyn = gdDBS.SelectBankMaster("DISTINCT DAKJNM", eBankRecordKubun.Bank, Trim(txtCABANK.Text), vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblBankName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))
    Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Shiten, Trim(txtCABANK.Text), Trim(txtCASITN.Text), vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"�x�X��_����" �œǂ߂Ȃ�
End Sub

Private Sub spnRireki_DownClick()
    '//��̃��R�[�h�Ɉړ�
    If True = gdDBS.MoveRecords(dbcHogoshaMaster, -1) Then '//�f�[�^�� DESC ORDER �������Ă���̂ł���ł悢
        On Error GoTo spnRireki_SpinDownError
        '//���Z�@�ւ̖��̂�\��
        Call SetBankAndShiten
'//�ŏI�̃f�[�^�̂ݕҏW�\�Ƃ���
        If dbcHogoshaMaster.Recordset.IsFirst Then
            If eShoriKubun.Refer <> lblShoriKubun.Caption Then  '//�Q�ƈȊO�̎�
                dbcHogoshaMaster.Recordset.Edit     '//�����Ŕr�����|����
                Call pLockedControl(False)
                spnRireki.Visible = True
                '//���̃{�^���͎x�X���N���b�N�������Ɏg����悤�ɂ���.
                cmdKakutei.Enabled = False
            Else
                Me.txtCAKYCD.Enabled = True
                Me.txtCAKSCD.Enabled = True
                Me.txtCAHGCD.Enabled = True
                Me.txtCAHGCD.SetFocus
                cmdUpdate.Enabled = False
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
        '//���Z�@�ւ̖��̂�\��
        Call SetBankAndShiten
'//�ŏI�̃f�[�^�̂ݕҏW�\�Ƃ���
'        dbcKeiyakushaMaster.Recordset.Edit
        Call mForm.LockedControlAllTextBox
        cmdEnd.Enabled = True
        cmdCancel.Enabled = True
    Else
        Call MsgBox("����ȑO�Ƀf�[�^�͂���܂���.", vbInformation, mCaption)
    End If
End Sub

Private Sub txtCABANK_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCABANK_LostFocus()
    If 0 <= Len(Trim(txtCABANK.Text)) And Len(Trim(txtCABANK.Text)) < 4 Then
        lblBankName.Caption = ""
        Exit Sub
    End If
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Set dyn = gdDBS.SelectBankMaster("DISTINCT DAKJNM", eBankRecordKubun.Bank, Trim(txtCABANK.Text), vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblBankName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))
End Sub

Private Sub txtCAFKxx_Change(Index As Integer)
    Call pButtonControl(True)
End Sub

Private Sub txtCAFKxx_DropOpen(Index As Integer, NoDefault As Boolean)
    txtCAFKxx(Index).Calendar.Holidays = gdDBS.Holiday(txtCAFKxx(Index).Year)
End Sub

Private Sub txtCAHKGK_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCAKJNM_Change()
    If Len(Trim(txtCAKJNM.Text)) = 0 Then
        txtCAKNNM.Text = ""
        txtCAKZNM.Text = ""
    End If
    Call pButtonControl(True)
End Sub

Private Sub txtCAKJNM_Furigana(Yomi As String)
'//2007/06/07 �J�i���ƌ������`�l��������
'    '//���݂̓ǂ݃J�i���ƌ������`�l���������Ȃ�ǂ݃J�i���ƌ������`�l���ɓ]��
'    If Trim(txtCAKNNM.Text) = Trim(txtCAKZNM.Text) Then
'        txtCAKNNM.Text = txtCAKNNM.Text & Yomi
'        txtCAKZNM.Text = txtCAKNNM.Text
'    Else
'        txtCAKNNM.Text = txtCAKNNM.Text & Yomi
'    End If
     txtCAKNNM.Text = txtCAKNNM.Text & Yomi
     txtCAKZNM.Text = txtCAKNNM.Text
End Sub

Private Sub txtCAHGCD_KeyDown(KeyCode As Integer, Shift As Integer)
    '// Return �܂��� Shift�{TAB �̂Ƃ��̂ݏ�������
    If Not (KeyCode = vbKeyReturn) Then
        Exit Sub
    End If
    
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
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
    txtCAHGCD.Text = Format(Val(txtCAHGCD.Text), "0000")
    sql = "SELECT * FROM tcHogoshaMaster"
    sql = sql & " WHERE CAITKB = '" & cboABKJNM.ItemData(cboABKJNM.ListIndex) & "'"
    sql = sql & "   AND CAKYCD = '" & txtCAKYCD.Text & "'"
    sql = sql & "   AND CAKSCD = '" & txtCAKSCD.Text & "'"
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
    mIsActivated = False    '//���R�[�h�\�����̃C�x���g���E��Ȃ��悤�Ƀt���O��ݒ�
    '//��񃁃b�Z�[�W�}�~
    dbcHogoshaMaster.RecordSource = sql
    Call dbcHogoshaMaster.Refresh
    On Error GoTo txtCAHGCD_KeyDownError        '//�r������p�G���[�g���b�v
    If 0& = dbcHogoshaMaster.Recordset.RecordCount Then
        '//�V�K�o�^
        dbcHogoshaMaster.Recordset.AddNew
        lblCAITKB.Caption = cboABKJNM.ItemData(cboABKJNM.ListIndex)
        lblCAKYCD.Caption = txtCAKYCD.Text
        lblCAKSCD.Caption = txtCAKSCD.Text
        lblCAHGCD.Caption = txtCAHGCD.Text
        lblCASQNO.Caption = gdDBS.sysDate("yyyymmdd")
        lblCAKKBN.Caption = 0
        lblCAKZSB.Caption = 1
        txtCAKYxx(0).Number = 20000101 '//��U�l��ݒ肵�Ȃ��Ɓu�O�v���Z�b�g����Ȃ��F�s�v�c�H
        txtCAKYxx(0).Number = 0
        txtCAKYxx(1).Number = gdDBS.LastDay(0)
        txtCAFKxx(0).Number = 20000101 '//��U�l��ݒ肵�Ȃ��Ɓu�O�v���Z�b�g����Ȃ��F�s�v�c�H
        txtCAFKxx(0).Number = 0
        txtCAFKxx(1).Number = gdDBS.LastDay(0)
    Else
        '//2007/06/06   ��s���E�x�X���̓ǂݍ��݂������ł���悤�ɕύX
        '//             �Ǎ��ݎ��� Change()=���̕\�� �C�x���g���Ԃ� �x�X�R�[�h�E��s�R�[�h�̏��ɂȂ�x�X�����\������Ȃ����Ƃ�����
        If eBankKubun.KinnyuuKikan = dbcHogoshaMaster.Recordset.Fields("CAKKBN").Value Then
            Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Bank, _
               dbcHogoshaMaster.Recordset.Fields("CABANK").Value, vDate:=gdDBS.sysDate("YYYYMMDD"))
            lblBankName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))
            Set dyn = Nothing
            Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Shiten, _
                dbcHogoshaMaster.Recordset.Fields("CABANK").Value, _
                dbcHogoshaMaster.Recordset.Fields("CASITN").Value, vDate:=gdDBS.sysDate("YYYYMMDD"))
            lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"�x�X��_����" �œǂ߂Ȃ�
            Set dyn = Nothing
        End If
        '//�C���E�폜
        Call dbcHogoshaMaster.Recordset.MoveFirst
        Call dbcHogoshaMaster.Recordset.Edit
'        Call dbcHogoshaMaster.UpdateRecord
    End If
    '//�Q�ƂŖ�����΃{�^���̐���J�n
    If False = optShoriKubun(eShoriKubun.Refer).Value Then
        Call pLockedControl(False)
    End If
    spnRireki.Visible = dbcHogoshaMaster.Recordset.RecordCount > 1
    '//���̃{�^���͎x�X���N���b�N�������Ɏg����悤�ɂ���.
    cmdKakutei.Enabled = False
    '//��񃁃b�Z�[�W�}�~
    '//�R���g���[����ی�ҁi�����j�ɂ����������߂ɂ��܂��Ȃ��F���ɕ��@��������Ȃ��H
    'If True = optShoriKubun(eShoriKubun.Refer).Value Then
        Call SendKeys("+{TAB}")
    'Else
    '    Call SendKeys("+{TAB}+{TAB}")
    'End If
    '//���~�{�^���͎Q�ƈȊO�͂��ł������\�ɁI
    Call pButtonControl(optShoriKubun(eShoriKubun.Delete).Value, True)
    '//���~�{�^���͂��ł������\�ɁI
    If Not optShoriKubun(eShoriKubun.Refer).Value Then
        cmdCancel.Visible = True
        cmdCancel.Enabled = True
    End If
    Exit Sub
txtCAHGCD_KeyDownError:
    Call gdDBS.ErrorCheck(dbcHogoshaMaster.Database)    '//�r������p�G���[�g���b�v
End Sub

Private Sub txtCAFKxx_LostFocus(Index As Integer)
    lblCAFKxx(Index).Caption = Val(gdDBS.Nz(txtCAFKxx(Index).Number))
End Sub

Private Sub txtCAKNNM_Change()
    txtCAKZNM.Text = txtCAKNNM.Text '//2007/06/07 �ی�Җ�(�J�i)���������`�l��
    Call pButtonControl(True)
End Sub

Private Sub txtCAKSCD_LostFocus()
'//2006/04/26 �O�[�����ߍ���
    txtCAKSCD.Text = Format(Val(txtCAKSCD.Text), "000")
End Sub

Private Sub txtCAKYxx_Change(Index As Integer)
    Call pButtonControl(True)
End Sub

Private Sub txtCAKYxx_DropOpen(Index As Integer, NoDefault As Boolean)
    txtCAKYxx(Index).Calendar.Holidays = gdDBS.Holiday(txtCAKYxx(Index).Year)
End Sub

Private Sub txtCAKYxx_LostFocus(Index As Integer)
    lblCAKYxx(Index).Caption = Val(gdDBS.Nz(txtCAKYxx(Index).Number))
End Sub

Private Sub txtCAKYCD_KeyDown(KeyCode As Integer, Shift As Integer)
    '// Return �܂��� Shift�{TAB �̂Ƃ��̂ݏ�������
    If Not (KeyCode = vbKeyReturn) Then
        Exit Sub
    End If
    
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim msg As String
        
    If "" = Trim(txtCAKYCD.Text) Then
        Exit Sub
    End If
'//2006/04/26 �O�[�����ߍ���
    txtCAKYCD.Text = Format(Val(txtCAKYCD.Text), "00000")
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//    sql = "SELECT DISTINCT BAITKB,BAKYCD,BAKSCD,BAKJNM FROM tbKeiyakushaMaster"
    sql = "SELECT DISTINCT BAITKB,BAKYCD,BAKJNM FROM tbKeiyakushaMaster"
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
        Call MsgBox("�_��҂�����ԁA�������͊Y���f�[�^�����݂��܂���.(" & lblKeiyakushaCode.Caption & ")", vbInformation, mCaption)
        Call txtCAKYCD.SetFocus
        Exit Sub
    End If
    lblBAKJNM.Caption = dyn.Fields("BAKJNM")
#If 0 Then
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
    Call cboCAKSCDz.Clear
    Do Until dyn.EOF
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//        Call cboCAKSCDz.AddItem(dyn.Fields("BAKSCD"))
        Call dyn.MoveNext
    Loop
    cboCAKSCDz.ListIndex = 0
#End If
    Call dyn.Close
End Sub

Private Function pUpdateErrorCheck() As Boolean
'//2006/06/26 �U���݈˗����ɂ��������W�b�N���L��̂Œ���
    '///////////////////////////////
    '//�K�{���͍��ڂƐ������`�F�b�N
    
    Dim str As New StringClass
    Dim obj As Object, msg As String
    '//�ی�ҁE�������͕̂K�{
    If txtCAKJNM.Text = "" Then
        Set obj = txtCAKJNM
        msg = "�ی�Җ�(����)�͕K�{���͂ł�."
    ElseIf False = str.CheckLength(txtCAKJNM.Text) Then
        Set obj = txtCAKJNM
        msg = "�ی�Җ�(����)�ɔ��p���܂܂�Ă��܂�."
    End If
    '//�ی�ҁE�J�i���͕̂K�{
    '//2007/06/07 �K�{ �����F�������`�l�Ɠ����l�Ƃ����
    If txtCAKNNM.Text = "" Then
        Set obj = txtCAKNNM
        msg = "�ی�Җ�(�J�i)�͕K�{���͂ł�."
    ElseIf False = str.CheckLength(txtCAKNNM.Text, vbNarrow) Then
        Set obj = txtCAKNNM
        msg = "�ی�Җ�(�J�i)�ɑS�p���܂܂�Ă��܂�."
    ElseIf 0 < InStr(txtCAKNNM.Text, "�") Then
        Set obj = txtCAKNNM
        msg = "�ی�Җ�(�J�i)�ɒ������܂܂�Ă��܂�."
    End If
    If IsNull(txtCAKYxx(1).Number) Then
        Set obj = txtCAKYxx(1)
        msg = "�_����Ԃ̏I�����͕K�{���͂ł�."
    ElseIf txtCAKYxx(0).Text > txtCAKYxx(1).Text Then
        Set obj = txtCAKYxx(0)
        msg = "�_����Ԃ��s���ł�."
    ElseIf IsNull(txtCAFKxx(1).Number) Then
        Set obj = txtCAFKxx(1)
        msg = "�U�֊��Ԃ̏I�����͕K�{���͂ł�."
    ElseIf txtCAFKxx(0).Text > txtCAFKxx(1).Text Then
        Set obj = txtCAFKxx(0)
        msg = "�U�֊��Ԃ��s���ł�."
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
        ElseIf txtCAKZNO.Text = "" Then
            Set obj = txtCAKZNO
            msg = "�����ԍ��͕K�{���͂ł�."
        End If
    ElseIf lblCAKKBN.Caption = eBankKubun.YuubinKyoku Then
        If txtCAYBTK.Text = "" Then
            Set obj = txtCAYBTK
            msg = "�ʒ��L���͕K�{���͂ł�."
        ElseIf txtCAYBTN.Text = "" Then
            Set obj = txtCAYBTN
            msg = "�ʒ��ԍ��͕K�{���͂ł�."
        ElseIf "1" <> Right(txtCAYBTN.Text, 1) Then
'//2006/04/26 �����ԍ��`�F�b�N
            Set obj = txtCAYBTN
            msg = "�ʒ��ԍ��̖������u�P�v�ȊO�ł�."
        End If
    End If
    '//2007/06/07 �K�{ �����F�������`�l�Ɠ����l�Ƃ����
'    If txtCAKZNM.Text = "" Then
'        Set obj = txtCAKZNM
'        msg = "�������`�l(�J�i)�͕K�{���͂ł�."
'    End If
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
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim AddRireki As String
    
    sql = "SELECT * FROM tcHogoshaMaster"
    sql = sql & " WHERE CAITKB = '" & lblCAITKB.Caption & "'"
    sql = sql & "   AND CAKYCD = '" & lblCAKYCD.Caption & "'"
    sql = sql & "   AND CAKSCD = '" & lblCAKSCD.Caption & "'"
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
            AddRireki = "���ԋ@��"
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
                              "�u�㏑���v���݂̃f�[�^�ɏ㏑������ꍇ�͂��̃{�^���������܂�."
        .lblFurikomi.Caption = "�U�֊J�n��"
        Call .Show(vbModal)
        '//���j������邩�킩��Ȃ��̂Ń��[�J���R�s�[���Ă���
        Dim PushButton As Integer, KeiyakuEnd As Long, FurikaeEnd As Long
        PushButton = .mPushButton
        KeiyakuEnd = .mKeiyakuEnd
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
    sql = sql & "   AND CAKSCD = '" & lblCAKSCD.Caption & "'"
    sql = sql & "   AND CAHGCD = '" & lblCAHGCD.Caption & "'"
    sql = sql & "   AND CASQNO = -1"
    Call gdDBS.Database.ExecuteSQL(sql)
    
    '////////////////////////////////////////////////
    '//�e�[�u����`���ύX���ꂽ�ꍇ���ӂ��邱�ƁI�I
    '//2007/06/11 �x���肵���ڒǉ��F���܂�^�p�サ�Ă��Ȃ��H
    Dim FixedCol As String
    FixedCol = "CAITKB,CAKYCD,CAKSCD,CAKJNM,CAHGCD,CAKNNM," & _
               "CASTNM,CAKKBN,CABANK,CASITN,CAKZSB,CAKZNO," & _
               "CAKZNM,CAYBTK,CAYBTN,CAKYST,CAFKST,CASKGK," & _
               "CAHKGK,CAKYDT,CAKYFG,CATRFG,CAADDT,CAUSID," & _
               "CANWDT,CAKYSR,CACHEK"
                
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
    sql = sql & "   AND CAKSCD = '" & lblCAKSCD.Caption & "'"
    sql = sql & "   AND CAHGCD = '" & lblCAHGCD.Caption & "'"
    sql = sql & "   AND CASQNO =  " & lblCASQNO.Caption
    Call gdDBS.Database.ExecuteSQL(sql)
    
    txtCAKYxx(0).Number = KeiyakuEnd
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
    sql = sql & "   AND CAKSCD = '" & lblCAKSCD.Caption & "'"
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
    sql = sql & "   AND CAKSCD = '" & lblCAKSCD.Caption & "'"
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

Private Sub txtCAKZNM_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCAKZNO_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCASITN_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCASITN_LostFocus()
    If 0 <= Len(Trim(txtCASITN.Text)) And Len(Trim(txtCASITN.Text)) < 3 Then
        lblShitenName.Caption = ""
        Exit Sub
    End If
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Shiten, Trim(txtCABANK.Text), Trim(txtCASITN.Text), vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"�x�X��_����" �œǂ߂Ȃ�
End Sub

Private Sub txtCASKGK_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCASTNM_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCAYBTK_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCAYBTK_LostFocus()
'//2006/04/26 �O�[�����ߍ���
    If "" <> txtCAYBTK.Text Then
        txtCAYBTK.Text = Format(Val(txtCAYBTK.Text), "000")
    End If
End Sub

Private Sub txtCAYBTN_Change()
    Call pButtonControl(True)
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
