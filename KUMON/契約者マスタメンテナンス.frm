VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmKeiyakushaMaster 
   Caption         =   "�_��҃}�X�^�����e�i���X"
   ClientHeight    =   7365
   ClientLeft      =   1155
   ClientTop       =   2820
   ClientWidth     =   11445
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11445
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
      ItemData        =   "�_��҃}�X�^�����e�i���X.frx":0000
      Left            =   1680
      List            =   "�_��҃}�X�^�����e�i���X.frx":000D
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "InputKey"
      Top             =   840
      Width           =   1755
   End
   Begin VB.Frame fraKinnyuuKikan 
      Caption         =   "�U������"
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
      Left            =   6480
      TabIndex        =   30
      Top             =   300
      Width           =   4635
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
         TabIndex        =   54
         Top             =   420
         Width           =   3855
         Begin imText6Ctl.imText txtBAKZNO 
            DataField       =   "BAKZNO"
            DataSource      =   "dbcKeiyakushaMaster"
            Height          =   285
            Left            =   1140
            TabIndex        =   39
            Top             =   1380
            Width           =   795
            _Version        =   65537
            _ExtentX        =   1402
            _ExtentY        =   503
            Caption         =   "�_��҃}�X�^�����e�i���X.frx":002B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�_��҃}�X�^�����e�i���X.frx":0097
            Key             =   "�_��҃}�X�^�����e�i���X.frx":00B5
            MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":00F9
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
         Begin imText6Ctl.imText txtBASITN 
            DataField       =   "BASITN"
            DataSource      =   "dbcKeiyakushaMaster"
            Height          =   285
            Left            =   1200
            TabIndex        =   34
            Top             =   660
            Width           =   375
            _Version        =   65537
            _ExtentX        =   661
            _ExtentY        =   503
            Caption         =   "�_��҃}�X�^�����e�i���X.frx":0115
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�_��҃}�X�^�����e�i���X.frx":0181
            Key             =   "�_��҃}�X�^�����e�i���X.frx":019F
            MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":01E3
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
         Begin imText6Ctl.imText txtBABANK 
            DataField       =   "BABANK"
            DataSource      =   "dbcKeiyakushaMaster"
            Height          =   285
            Left            =   1200
            TabIndex        =   33
            Top             =   300
            Width           =   495
            _Version        =   65537
            _ExtentX        =   873
            _ExtentY        =   503
            Caption         =   "�_��҃}�X�^�����e�i���X.frx":01FF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�_��҃}�X�^�����e�i���X.frx":026B
            Key             =   "�_��҃}�X�^�����e�i���X.frx":0289
            MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":02CD
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
            TabIndex        =   35
            Top             =   900
            Width           =   2535
            Begin VB.OptionButton optBAKZSB 
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
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   480
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.OptionButton optBAKZSB 
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
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   180
               Width           =   675
            End
            Begin VB.OptionButton optBAKZSB 
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
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   180
               Width           =   675
            End
            Begin VB.Label lblBAKZSB 
               BackColor       =   &H000000FF&
               Caption         =   "�������"
               DataField       =   "BAKZSB"
               DataSource      =   "dbcKeiyakushaMaster"
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
               TabIndex        =   78
               Top             =   180
               Width           =   795
            End
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
            TabIndex        =   55
            Top             =   660
            Width           =   1875
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
            TabIndex        =   60
            Top             =   1380
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
            TabIndex        =   59
            Top             =   1020
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
            TabIndex        =   58
            Top             =   660
            Width           =   795
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
            TabIndex        =   57
            Top             =   300
            Width           =   795
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
            TabIndex        =   56
            Top             =   300
            Width           =   1875
         End
      End
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
         TabIndex        =   61
         Top             =   1020
         Width           =   4035
         Begin imText6Ctl.imText txtBAYBTK 
            DataField       =   "BAYBTK"
            DataSource      =   "dbcKeiyakushaMaster"
            Height          =   285
            Left            =   1860
            TabIndex        =   40
            Top             =   480
            Width           =   375
            _Version        =   65537
            _ExtentX        =   661
            _ExtentY        =   503
            Caption         =   "�_��҃}�X�^�����e�i���X.frx":02E9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�_��҃}�X�^�����e�i���X.frx":0355
            Key             =   "�_��҃}�X�^�����e�i���X.frx":0373
            MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":03B7
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
         Begin imText6Ctl.imText txtBAYBTN 
            DataField       =   "BAYBTN"
            DataSource      =   "dbcKeiyakushaMaster"
            Height          =   285
            Left            =   1860
            TabIndex        =   41
            Top             =   960
            Width           =   855
            _Version        =   65537
            _ExtentX        =   1508
            _ExtentY        =   503
            Caption         =   "�_��҃}�X�^�����e�i���X.frx":03D3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�_��҃}�X�^�����e�i���X.frx":043F
            Key             =   "�_��҃}�X�^�����e�i���X.frx":045D
            MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":04A1
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
            TabIndex        =   63
            Top             =   480
            Width           =   1275
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
            TabIndex        =   62
            Top             =   960
            Width           =   1275
         End
      End
      Begin VB.OptionButton optBAKKBN 
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
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   1395
      End
      Begin VB.OptionButton optBAKKBN 
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
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin imText6Ctl.imText txtBAKZNM 
         DataField       =   "BAKZNM"
         DataSource      =   "dbcKeiyakushaMaster"
         Height          =   285
         Left            =   540
         TabIndex        =   42
         Top             =   2580
         Width           =   3735
         _Version        =   65537
         _ExtentX        =   6588
         _ExtentY        =   503
         Caption         =   "�_��҃}�X�^�����e�i���X.frx":04BD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "�_��҃}�X�^�����e�i���X.frx":0529
         Key             =   "�_��҃}�X�^�����e�i���X.frx":0547
         MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":058B
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
         Caption         =   "�������`�l��(�J�i)"
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
         Left            =   600
         TabIndex        =   95
         Top             =   2340
         Width           =   1575
      End
      Begin VB.Label lblBAKKBN 
         BackColor       =   &H000000FF&
         Caption         =   "���Z�@�֎��"
         DataField       =   "BAKKBN"
         DataSource      =   "dbcKeiyakushaMaster"
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
         TabIndex        =   81
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkBAKYFG 
      Caption         =   "���"
      DataField       =   "BAKYFG"
      Height          =   315
      Left            =   4140
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5160
      Width           =   675
   End
   Begin imDate6Ctl.imDate txtBAKYxx 
      DataField       =   "BAKYST"
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   24
      Top             =   5160
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   556
      Calendar        =   "�_��҃}�X�^�����e�i���X.frx":05A7
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":0727
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":0793
      Keys            =   "�_��҃}�X�^�����e�i���X.frx":07B1
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":080F
      Spin            =   "�_��҃}�X�^�����e�i���X.frx":082B
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
   Begin imText6Ctl.imText txtBATELE 
      DataField       =   "BATELE"
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   1680
      TabIndex        =   19
      Top             =   4080
      Width           =   1395
      _Version        =   65537
      _ExtentX        =   2461
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":0853
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":08BF
      Key             =   "�_��҃}�X�^�����e�i���X.frx":08DD
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":0921
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
      Text            =   ""
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
   Begin imText6Ctl.imText txtBAADJ1 
      DataField       =   "BAADJ1"
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   1680
      TabIndex        =   16
      Top             =   3000
      Width           =   4635
      _Version        =   65537
      _ExtentX        =   8176
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":093D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":09A9
      Key             =   "�_��҃}�X�^�����e�i���X.frx":09C7
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":0A0B
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
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   50
      LengthAsByte    =   -1
      Text            =   "�Z�������P�D�D�D�D�D�D�D�D�D�D�D�D�D�D�D�D�D�D�D��"
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
   Begin imText6Ctl.imText txtBAZPC1 
      DataField       =   "BAZPC1"
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   1680
      TabIndex        =   14
      Top             =   2640
      Width           =   375
      _Version        =   65537
      _ExtentX        =   661
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":0A27
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":0A93
      Key             =   "�_��҃}�X�^�����e�i���X.frx":0AB1
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":0AF5
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
   Begin imText6Ctl.imText txtBAKJNM 
      DataField       =   "BAKJNM"
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   1920
      Width           =   3735
      _Version        =   65537
      _ExtentX        =   6588
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":0B11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":0B7D
      Key             =   "�_��҃}�X�^�����e�i���X.frx":0B9B
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":0BDF
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
      Text            =   "���������D�D�D�D�D�D�D�D�D�D�D�D�D�D�D��"
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
   Begin imText6Ctl.imText txtBAKYCD 
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Tag             =   "InputKey"
      Top             =   1200
      Width           =   615
      _Version        =   65537
      _ExtentX        =   1085
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":0BFB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":0C67
      Key             =   "�_��҃}�X�^�����e�i���X.frx":0C85
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":0CC9
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "���~(&C)"
      Height          =   435
      Left            =   2400
      TabIndex        =   80
      Top             =   6720
      Width           =   1395
   End
   Begin ORADCLibCtl.ORADC dbcKeiyakushaMaster 
      Height          =   315
      Left            =   6840
      Top             =   7020
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
      DatabaseName    =   "dcssvr03"
      Connect         =   "kumon/kumon"
      RecordSource    =   "SELECT * FROM tbKeiyakushaMaster"
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�X�V(&U)"
      Height          =   435
      Left            =   480
      TabIndex        =   72
      Top             =   6720
      Width           =   1395
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
      Left            =   6360
      TabIndex        =   64
      Top             =   3300
      Width           =   4875
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
         RecordSource    =   ""
      End
      Begin VB.CommandButton cmdKakutei 
         Caption         =   "�m��(&K)"
         Height          =   375
         Left            =   3660
         TabIndex        =   67
         Top             =   2700
         Width           =   975
      End
      Begin VB.ComboBox cboShitenYomi 
         Height          =   300
         ItemData        =   "�_��҃}�X�^�����e�i���X.frx":0CE5
         Left            =   3900
         List            =   "�_��҃}�X�^�����e�i���X.frx":0D0A
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   180
         Width           =   855
      End
      Begin VB.ComboBox cboBankYomi 
         Height          =   300
         ItemData        =   "�_��҃}�X�^�����e�i���X.frx":0D4C
         Left            =   1500
         List            =   "�_��҃}�X�^�����e�i���X.frx":0D71
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   180
         Width           =   855
      End
      Begin MSDBCtls.DBList dblBankList 
         Bindings        =   "�_��҃}�X�^�����e�i���X.frx":0DB3
         Height          =   2040
         Left            =   120
         TabIndex        =   68
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
         Bindings        =   "�_��҃}�X�^�����e�i���X.frx":0DC9
         Height          =   2040
         Left            =   2400
         TabIndex        =   69
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
         TabIndex        =   71
         Top             =   240
         Width           =   1395
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
         TabIndex        =   70
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraShoriKubun 
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
      Left            =   420
      TabIndex        =   0
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
         Left            =   2820
         TabIndex        =   99
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
         TabIndex        =   1
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         Left            =   1560
         TabIndex        =   79
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "�I��(&X)"
      Height          =   495
      Left            =   9600
      TabIndex        =   53
      Top             =   6720
      Width           =   1335
   End
   Begin imText6Ctl.imText txtBAKNNM 
      DataField       =   "BAKNNM"
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Top             =   2280
      Width           =   3735
      _Version        =   65537
      _ExtentX        =   6588
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":0DE1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":0E4D
      Key             =   "�_��҃}�X�^�����e�i���X.frx":0E6B
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":0EAF
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
      Text            =   "Ŷ�Ҳ..................................*"
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
   Begin imText6Ctl.imText txtBAADJ2 
      DataField       =   "BAADJ2"
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   1680
      TabIndex        =   17
      Top             =   3360
      Width           =   4635
      _Version        =   65537
      _ExtentX        =   8176
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":0ECB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":0F37
      Key             =   "�_��҃}�X�^�����e�i���X.frx":0F55
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":0F99
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
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   50
      LengthAsByte    =   -1
      Text            =   "�����n�C�c�R�S�T��"
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
   Begin imText6Ctl.imText txtBAADJ3 
      DataField       =   "BAADJ3"
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   1680
      TabIndex        =   18
      Top             =   3720
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":0FB5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":1021
      Key             =   "�_��҃}�X�^�����e�i���X.frx":103F
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":1083
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
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   30
      LengthAsByte    =   -1
      Text            =   "��ؕ��S�T�U�V�W�X�O�P�Q�R�S�T"
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
   Begin imText6Ctl.imText txtBAFAXI 
      DataField       =   "BAFAXI"
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   1680
      TabIndex        =   22
      Top             =   4800
      Width           =   1395
      _Version        =   65537
      _ExtentX        =   2461
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":109F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":110B
      Key             =   "�_��҃}�X�^�����e�i���X.frx":1129
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":116D
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
      Text            =   "06-6234-1235"
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
   Begin imText6Ctl.imText txtBAKKRN 
      DataField       =   "BAKKRN"
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   1680
      TabIndex        =   21
      Top             =   4440
      Width           =   1395
      _Version        =   65537
      _ExtentX        =   2461
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":1189
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":11F5
      Key             =   "�_��҃}�X�^�����e�i���X.frx":1213
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":1257
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
      Text            =   "090-010-1234"
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
   Begin imDate6Ctl.imDate txtBAKYxx 
      DataField       =   "BAKYED"
      Height          =   315
      Index           =   1
      Left            =   3000
      TabIndex        =   25
      Top             =   5160
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   556
      Calendar        =   "�_��҃}�X�^�����e�i���X.frx":1273
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":13F3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":145F
      Keys            =   "�_��҃}�X�^�����e�i���X.frx":147D
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":14DB
      Spin            =   "�_��҃}�X�^�����e�i���X.frx":14F7
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
   Begin imDate6Ctl.imDate txtBAFKxx 
      DataField       =   "BAFKST"
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   27
      Top             =   5580
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   556
      Calendar        =   "�_��҃}�X�^�����e�i���X.frx":151F
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":169F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":170B
      Keys            =   "�_��҃}�X�^�����e�i���X.frx":1729
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":1787
      Spin            =   "�_��҃}�X�^�����e�i���X.frx":17A3
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
   Begin imDate6Ctl.imDate txtBAFKxx 
      DataField       =   "BAFKED"
      Height          =   315
      Index           =   1
      Left            =   3000
      TabIndex        =   28
      Top             =   5580
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   556
      Calendar        =   "�_��҃}�X�^�����e�i���X.frx":17CB
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":194B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":19B7
      Keys            =   "�_��҃}�X�^�����e�i���X.frx":19D5
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":1A33
      Spin            =   "�_��҃}�X�^�����e�i���X.frx":1A4F
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
   Begin imText6Ctl.imText txtBAKSCDz 
      DataField       =   "BAKSCD"
      Height          =   285
      Left            =   5880
      TabIndex        =   8
      Tag             =   "InputKey"
      Top             =   900
      Visible         =   0   'False
      Width           =   435
      _Version        =   65537
      _ExtentX        =   767
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":1A77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":1AE3
      Key             =   "�_��҃}�X�^�����e�i���X.frx":1B01
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":1B45
      BackColor       =   255
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
   Begin imText6Ctl.imText txtBATELJ 
      DataField       =   "BATELJ"
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   3780
      TabIndex        =   20
      Top             =   4080
      Width           =   1395
      _Version        =   65537
      _ExtentX        =   2461
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":1B61
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":1BCD
      Key             =   "�_��҃}�X�^�����e�i���X.frx":1BEB
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":1C2F
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
      Text            =   ""
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
   Begin imText6Ctl.imText txtBAFAXJ 
      DataField       =   "BAFAXJ"
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   3780
      TabIndex        =   23
      Top             =   4800
      Width           =   1395
      _Version        =   65537
      _ExtentX        =   2461
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":1C4B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":1CB7
      Key             =   "�_��҃}�X�^�����e�i���X.frx":1CD5
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":1D19
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
      Text            =   "06-6234-1235"
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
      Left            =   6840
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
   Begin imText6Ctl.imText txtBAZPC2 
      DataField       =   "BAZPC2"
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Top             =   2640
      Width           =   495
      _Version        =   65537
      _ExtentX        =   873
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":1D35
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":1DA1
      Key             =   "�_��҃}�X�^�����e�i���X.frx":1DBF
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":1E03
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
   Begin imText6Ctl.imText txtBAKSNO 
      DataField       =   "BAKSNO"
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   1560
      Width           =   735
      _Version        =   65537
      _ExtentX        =   1296
      _ExtentY        =   503
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":1E1F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":1E8B
      Key             =   "�_��҃}�X�^�����e�i���X.frx":1EA9
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":1EED
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
   Begin imNumber6Ctl.imNumber ImNumber2 
      DataField       =   "BASOFU"
      DataSource      =   "dbcKeiyakushaMaster"
      Height          =   285
      Left            =   1680
      TabIndex        =   29
      Top             =   6000
      Visible         =   0   'False
      Width           =   675
      _Version        =   65537
      _ExtentX        =   1191
      _ExtentY        =   503
      Calculator      =   "�_��҃}�X�^�����e�i���X.frx":1F09
      Caption         =   "�_��҃}�X�^�����e�i���X.frx":1F29
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�����e�i���X.frx":1F95
      Keys            =   "�_��҃}�X�^�����e�i���X.frx":1FB3
      MouseIcon       =   "�_��҃}�X�^�����e�i���X.frx":1FFD
      Spin            =   "�_��҃}�X�^�����e�i���X.frx":2019
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   255
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSComCtl2.UpDown spnRireki 
      Height          =   435
      Left            =   2490
      TabIndex        =   98
      ToolTipText     =   "�O��̗����Ɉړ�"
      Top             =   1185
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   767
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label Label30 
      Alignment       =   1  '�E����
      BackColor       =   &H000000FF&
      Caption         =   "���t��]����"
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
      TabIndex        =   97
      Top             =   6000
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label29 
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
      Left            =   240
      TabIndex        =   96
      Top             =   1590
      Width           =   1275
   End
   Begin VB.Label Label27 
      Caption         =   "�|"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   94
      Top             =   2655
      Width           =   195
   End
   Begin VB.Label lblBAADDT 
      BackColor       =   &H000000FF&
      Caption         =   "�쐬��"
      DataField       =   "BAADDT"
      DataSource      =   "dbcKeiyakushaMaster"
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
      TabIndex        =   93
      Top             =   6780
      Width           =   1875
   End
   Begin VB.Label lblBAKYxx 
      BackColor       =   &H000000FF&
      Caption         =   "�_��J�n��"
      DataField       =   "BAKYST"
      DataSource      =   "dbcKeiyakushaMaster"
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
      Left            =   4200
      TabIndex        =   92
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblBAKYxx 
      BackColor       =   &H000000FF&
      Caption         =   "�_��I����"
      DataField       =   "BAKYED"
      DataSource      =   "dbcKeiyakushaMaster"
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
      Left            =   5220
      TabIndex        =   91
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblBAFKxx 
      BackColor       =   &H000000FF&
      Caption         =   "�U���J�n��"
      DataField       =   "BAFKST"
      DataSource      =   "dbcKeiyakushaMaster"
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
      Left            =   4200
      TabIndex        =   90
      Top             =   5940
      Width           =   975
   End
   Begin VB.Label lblBAFKxx 
      BackColor       =   &H000000FF&
      Caption         =   "�U���I����"
      DataField       =   "BAFKED"
      DataSource      =   "dbcKeiyakushaMaster"
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
      Left            =   5220
      TabIndex        =   89
      Top             =   5940
      Width           =   975
   End
   Begin VB.Label lblBAKYFG 
      BackColor       =   &H000000FF&
      Caption         =   "���t���O"
      DataField       =   "BAKYFG"
      DataSource      =   "dbcKeiyakushaMaster"
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
      Left            =   5040
      TabIndex        =   88
      Top             =   5220
      Width           =   375
   End
   Begin VB.Label lblBAUPDT 
      BackColor       =   &H000000FF&
      Caption         =   "�X�V��"
      DataField       =   "BAUPDT"
      DataSource      =   "dbcKeiyakushaMaster"
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
      TabIndex        =   87
      Top             =   7080
      Width           =   1875
   End
   Begin VB.Label lblBAUSID 
      BackColor       =   &H000000FF&
      Caption         =   "�X�V��"
      DataField       =   "BAUSID"
      DataSource      =   "dbcKeiyakushaMaster"
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
      TabIndex        =   86
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label lblBASQNO 
      BackColor       =   &H000000FF&
      Caption         =   "�_��҂r�d�p"
      DataField       =   "BASQNO"
      DataSource      =   "dbcKeiyakushaMaster"
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
      TabIndex        =   10
      Top             =   1500
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
      Left            =   240
      TabIndex        =   85
      Tag             =   "InputKey"
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label21 
      Alignment       =   1  '�E����
      Caption         =   "(����)"
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
      TabIndex        =   84
      Top             =   4830
      Width           =   495
   End
   Begin VB.Label Label20 
      Alignment       =   1  '�E����
      Caption         =   "(����)"
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
      TabIndex        =   83
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblBAKSCD 
      BackColor       =   &H000000FF&
      Caption         =   "�����敪"
      DataField       =   "BAKSCD"
      DataSource      =   "dbcKeiyakushaMaster"
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
      Left            =   5280
      TabIndex        =   9
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  '�E����
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
      Left            =   4980
      TabIndex        =   82
      Tag             =   "InputKey"
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label lblBAITKB 
      BackColor       =   &H000000FF&
      Caption         =   "�ϑ��ҋ敪"
      DataField       =   "BAITKB"
      DataSource      =   "dbcKeiyakushaMaster"
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
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblBAKYCD 
      BackColor       =   &H000000FF&
      Caption         =   "�_��Ҕԍ�"
      DataField       =   "BAKYCD"
      DataSource      =   "dbcKeiyakushaMaster"
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
      TabIndex        =   7
      Top             =   1200
      Width           =   975
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
      Left            =   9120
      TabIndex        =   77
      Top             =   60
      Width           =   1395
   End
   Begin VB.Label Label19 
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
      Left            =   240
      TabIndex        =   76
      Top             =   5220
      Width           =   1275
   End
   Begin VB.Label Label18 
      Alignment       =   1  '�E����
      Caption         =   "�U������"
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
      Top             =   5640
      Width           =   1275
   End
   Begin VB.Label Label17 
      Alignment       =   1  '�E����
      Caption         =   "�`"
      DataSource      =   "dbcKeiyakushaMaster"
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
      TabIndex        =   74
      Top             =   5220
      Width           =   255
   End
   Begin VB.Label Label16 
      Alignment       =   1  '�E����
      Caption         =   "�`"
      DataSource      =   "dbcKeiyakushaMaster"
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
      TabIndex        =   73
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   52
      Tag             =   "InputKey"
      Top             =   1200
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
      Caption         =   "�_��Җ�(����)"
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
      TabIndex        =   51
      Top             =   1965
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      Caption         =   "�_��Җ�(�J�i) "
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
      TabIndex        =   50
      Top             =   2310
      Width           =   1275
   End
   Begin VB.Label Label4 
      Alignment       =   1  '�E����
      Caption         =   "�X�֔ԍ�"
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
      TabIndex        =   49
      Top             =   2670
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   1  '�E����
      Caption         =   "�Z���P(����)"
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
      TabIndex        =   48
      Top             =   3015
      Width           =   1275
   End
   Begin VB.Label Label7 
      Alignment       =   1  '�E����
      Caption         =   "�Z���Q(����)"
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
      TabIndex        =   47
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label Label8 
      Alignment       =   1  '�E����
      Caption         =   "�Z���R(����)"
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
      TabIndex        =   46
      Top             =   3720
      Width           =   1275
   End
   Begin VB.Label Label9 
      Alignment       =   1  '�E����
      Caption         =   "�d�b�ԍ�(����)"
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
      TabIndex        =   45
      Top             =   4050
      Width           =   1275
   End
   Begin VB.Label Label10 
      Alignment       =   1  '�E����
      Caption         =   "�ً}�A����"
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
      TabIndex        =   44
      Top             =   4455
      Width           =   1275
   End
   Begin VB.Label Label11 
      Alignment       =   1  '�E����
      Caption         =   "FAX�ԍ�(����)"
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
      TabIndex        =   43
      Top             =   4800
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
Attribute VB_Name = "frmKeiyakushaMaster"
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
'    cmdEnd.Enabled = blMode
    spnRireki.Visible = False
    cmdKakutei.Enabled = Not blMode
End Sub

Private Sub chkBAKYFG_Click()
    lblBAKYFG.Caption = chkBAKYFG.Value
    Call pButtonControl(True)
End Sub

Private Sub chkBAKYFG_KeyDown(KeyCode As Integer, Shift As Integer)
    '//���t���O��ݒ肵���̂ŏI�����̓��͂𑣂�.
    '//KeyCode & Shift ���N���A���Ȃ��ƃo�b�t�@�Ɏc��H
    KeyCode = 0
    Shift = 0
    chkBAKYFG.Value = Choose(chkBAKYFG.Value + 1, 1, 0, 0)  '// Index=1,2,3
    Call MsgBox("���̕ύX�����m���܂����B" & vbCrLf & vbCrLf & "�_����ԋy�ѐU�֊��� �I�����̍Đݒ�����ĉ�����.", vbInformation + vbOKOnly, mCaption)
    Call txtBAKYxx(1).SetFocus
End Sub

Private Sub chkBAKYFG_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '//���t���O��ݒ肵���̂ŏI�����̓��͂𑣂�.
    If Button = vbLeftButton Then
        Call chkBAKYFG_KeyDown(vbKeySpace, 0)
    End If
End Sub

Private Sub lblBAKYFG_Change()
    chkBAKYFG.Value = Val(lblBAKYFG.Caption)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub pUpdateRecord()
'''//2002/10/18 ���̂܂܂̓��t�Ƃ���
'''    lblBAKYxx(0).Caption = gdDBS.FirstDay(txtBAKYxx(0).Number)
'''    lblBAKYxx(1).Caption = gdDBS.LastDay(txtBAKYxx(1).Number)
'''    lblBAFKxx(0).Caption = gdDBS.FirstDay(txtBAFKxx(0).Number)
'''    lblBAFKxx(1).Caption = gdDBS.LastDay(txtBAFKxx(1).Number)
    lblBAKYxx(0).Caption = Val(gdDBS.Nz(txtBAKYxx(0).Number))
    lblBAKYxx(1).Caption = Val(gdDBS.Nz(txtBAKYxx(1).Number))
    lblBAFKxx(0).Caption = Val(gdDBS.Nz(txtBAFKxx(0).Number))
    lblBAFKxx(1).Caption = Val(gdDBS.Nz(txtBAFKxx(1).Number))
'//2003/01/31 ���t���O�� NULL �ɂȂ�̂ŕύX
    lblBAKYFG.Caption = Val(chkBAKYFG.Value)
    lblBAUSID.Caption = gdDBS.LoginUserName
    If "" = lblBAADDT.Caption Then
        lblBAADDT.Caption = gdDBS.sysDate
    End If
    lblBAUPDT.Caption = gdDBS.sysDate
    Call dbcKeiyakushaMaster.UpdateRecord
'//2006/08/02 �_��҂̉�񎞂Ɍx����\��
    If 0 = chkBAKYFG.Value Then
        Exit Sub
    End If
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    sql = "SELECT COUNT(*) AS CNT FROM tcHogoshaMaster"
    sql = sql & " WHERE CAITKB = '" & lblBAITKB.Caption & "'"
    sql = sql & "   AND CAKYCD = '" & lblBAKYCD.Caption & "'"
'//2007/07/19 ��񂵂Ă��Ȃ��f�[�^����������悤�ɕύX
    sql = sql & "   AND NVL(CAKYFG,0) = 0 " & vbCrLf   '//�ی�҂͉���ԂłȂ��I
    sql = sql & "   AND CANWDT IS NULL "
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If dyn.Fields("CNT") Then
'//2007/07/19 ��񂵂Ă��Ȃ��f�[�^����������悤�ɕύX�����̂Ń��b�Z�[�W��ύX
'        Call MsgBox("�� �V�K�̕ی�Җ��͐V�K�̉��ی�҂� " & dyn.Fields("CNT") & " �����݂��܂��B" & vbCrLf & vbCrLf &
        Call MsgBox("�� �V�K�����̖����ی�҂� " & dyn.Fields("CNT") & " �����݂��܂��B" & vbCrLf & vbCrLf & _
                "�����U�ւ̐V�K�������s��v�ɂȂ�܂��B", vbInformation + vbOKOnly, Me.Caption)
    End If
    Call dyn.Close
End Sub

Private Sub cmdUpdate_Click()
    If lblShoriKubun.Caption = eShoriKubun.Delete Then
#If ORA_DEBUG = 1 Then
        Dim sql As String, dyn As OraDynaset
#Else
        Dim sql As String, dyn As Object
#End If
        sql = "SELECT COUNT(*) AS CNT FROM tcHogoshaMaster"
        sql = sql & " WHERE CAITKB = '" & lblBAITKB.Caption & "'"
        sql = sql & "   AND CAKYCD = '" & lblBAKYCD.Caption & "'"
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//        sql = sql & "   AND CAKSCD = '" & lblBAKSCD.Caption & "'"
'        sql = sql & "   AND CASQNO = '" & lblBASQNO.Caption & "'"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
        If Val(gdDBS.Nz(dyn.Fields("CNT"))) Then
            Call MsgBox("�ی�҃}�X�^�Ŏg�p����Ă��邽��" & vbCrLf & vbCrLf & "�폜���鎖�͏o���܂���.", vbCritical, mCaption)
            Exit Sub
        End If
        If vbOK <> MsgBox("�폜���܂����H" & vbCrLf & vbCrLf & "���ɖ߂����Ƃ͏o���܂���.", vbInformation + vbOKCancel + vbDefaultButton2, mCaption) Then
            Exit Sub
        Else
'//2002/11/26 OIP-00000 ORA-04108 �ŃG���[�ɂȂ�̂� Execute() �Ŏ��s����悤�ɕύX.
'// Oracle Data Control 8i(3.6) 9i(4.2) �̈Ⴂ���ȁH
'//            Call dbcKeiyakushaMaster.Recordset.Delete
            Call dbcKeiyakushaMaster.UpdateControls
            sql = "DELETE tbKeiyakushaMaster"
            sql = sql & " WHERE BAITKB = '" & lblBAITKB.Caption & "'"
            sql = sql & "   AND BAKYCD = '" & lblBAKYCD.Caption & "'"
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//            sql = sql & "   AND BAKSCD = '" & lblBAKSCD.Caption & "'"
            sql = sql & "   AND BASQNO =  " & lblBASQNO.Caption
            Call gdDBS.Database.ExecuteSQL(sql)
        End If
    Else
'//2013/02/26 �����ύX���̍X�V���̒ǉ��X�V�̍ۂɂQ�x pUpdateRecord() �����s�����̂𐧌䂷��
        mRirekiAddNewUpdate = False
        '//���͓��e�`�F�b�N�Ŏ���߂����̂ŏI��
        If False = pUpdateErrorCheck Then
            Exit Sub
        End If
        Call pUpdateRecord
    End If
    Call pLockedControl(True)
    Call txtBAKYCD.SetFocus ' cboABKJNM.SetFocus
    Call pButtonControl(False)
End Sub

Private Sub cmdCancel_Click()
    Call dbcKeiyakushaMaster.UpdateControls
    Call pLockedControl(True)
    Call txtBAKYCD.SetFocus ' cboABKJNM.SetFocus
    Call pButtonControl(False)
End Sub

Private Sub cmdEnd_Click()
    Call dbcKeiyakushaMaster.UpdateControls
    Unload Me
End Sub

Private Sub cmdKakutei_Click()
    If dblBankList.Text = "" Or dblShitenList.Text = "" Then
        Exit Sub
    End If
    txtBABANK.Text = Left(dblBankList.Text, 4)
    lblBankName.Caption = Mid(dblBankList.Text, 6)
    txtBASITN.Text = Left(dblShitenList.Text, 3)
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

Private Sub dblBankList_Click()
    cboShitenYomi.ListIndex = -1
    Call cboShitenYomi_Click
End Sub

Private Sub dblShitenList_Click()
    cmdKakutei.Enabled = dblBankList.Text <> ""
End Sub

Private Sub Form_Activate()
    If False = mIsActivated Then
        Call pButtonControl(False, True)
    End If
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    '//��s�ƗX�֋ǂ� Frame �𐮗񂷂�
    fraBank(1).Top = fraBank(0).Top
    fraBank(1).Left = fraBank(0).Left
    fraBank(1).Height = fraBank(0).Height
    fraBank(1).Width = fraBank(0).Width
'    fraBank(0).BackColor = Me.BackColor
'    fraBank(1).BackColor = Me.BackColor
    fraBank(0).BorderStyle = vbBSNone
    fraBank(1).BorderStyle = vbBSNone
    fraBankList.BorderStyle = vbBSNone
'    fraKouzaShubetsu.BackColor = Me.BackColor
    
    dbcBank.RecordSource = ""
    dbcShiten.RecordSource = ""
    dbcKeiyakushaMaster.RecordSource = ""
    dbcItakushaMaster.RecordSource = "SELECT * FROM taItakushaMaster ORDER BY ABITCD"
    dbcItakushaMaster.ReadOnly = True
    Call pLockedControl(True)
    Call mForm.pInitControl
    '//�ϑ��҃R�[�h���͎��͂��̒�`���O��
    'txtBAKYCD.KeyNext = ""
    'txtBAKSCD.KeyNext = ""
    '//�����l���Z�b�g�F�C�����[�h
    optShoriKubun(eShoriKubun.Refer).Value = True
    'Call txtBAITKB.SetFocus
    spnRireki.Visible = False
    lblBankName.Caption = ""
    lblShitenName.Caption = ""
    Call gdDBS.SetItakushaComboBox(cboABKJNM)
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmKeiyakushaMaster = Nothing
    Set mForm = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub lblBAKKBN_Change()
    optBAKKBN(Val(lblBAKKBN.Caption)).Value = True
End Sub

Private Sub lblBAFKxx_Change(Index As Integer)
    txtBAFKxx(Index).Number = Val(lblBAFKxx(Index).Caption)
End Sub

Private Sub lblBAKYxx_Change(Index As Integer)
    txtBAKYxx(Index).Number = Val(lblBAKYxx(Index).Caption)
End Sub

Private Sub lblBAKZSB_Change()
    optBAKZSB(Val(lblBAKZSB.Caption)).Value = True
End Sub

Private Sub optBAKKBN_Click(Index As Integer)
    fraKinnyuuKikan.Tag = Index
    Call fraBank(Index).ZOrder(0)
    fraBankList.Visible = Index = 0
    lblBAKKBN.Caption = Index
    '//�t�H�[�J�X��������̂Őݒ肷��.
    txtBABANK.TabStop = Index = eBankKubun.KinnyuuKikan
    txtBASITN.TabStop = Index = eBankKubun.KinnyuuKikan
    txtBAKZNO.TabStop = Index = eBankKubun.KinnyuuKikan
    txtBAYBTK.TabStop = Index = eBankKubun.YuubinKyoku
    txtBAYBTN.TabStop = Index = eBankKubun.YuubinKyoku
    Call pButtonControl(True)
End Sub

Private Sub optBAKZSB_Click(Index As Integer)
    lblBAKZSB.Caption = Index
    Call pButtonControl(True)
End Sub

Private Sub optShoriKubun_Click(Index As Integer)
    On Error Resume Next    'Form_Load()���Ƀt�H�[�J�X�𓖂Ă��Ȃ����G���[�ƂȂ�̂ŉ���̃G���[����
    lblShoriKubun.Caption = Index
    Call txtBAKYCD.SetFocus ' cboABKJNM.SetFocus
End Sub

Private Sub SetBankAndShiten()
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Set dyn = gdDBS.SelectBankMaster("DISTINCT DAKJNM", eBankRecordKubun.Bank, Trim(txtBABANK.Text), vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblBankName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))
    Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Shiten, Trim(txtBABANK.Text), Trim(txtBASITN.Text), vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"�x�X��_����" �œǂ߂Ȃ�
End Sub

Private Sub spnRireki_DownClick()
    If True = gdDBS.MoveRecords(dbcKeiyakushaMaster, -1) Then '//�f�[�^�� DESC ORDER �������Ă���̂ł���ł悢
        On Error GoTo spnRireki_SpinDownError
        '//���Z�@�ւ̖��̂�\��
        Call SetBankAndShiten
'//�ŏI�̃f�[�^�̂ݕҏW�\�Ƃ���
        If dbcKeiyakushaMaster.Recordset.IsFirst Then
            If eShoriKubun.Refer <> lblShoriKubun.Caption Then  '//�Q�ƈȊO�̎�
                dbcKeiyakushaMaster.Recordset.Edit      '//�����Ŕr�����|����
                Call pLockedControl(False)
                spnRireki.Visible = True
                '//���̃{�^���͎x�X���N���b�N�������Ɏg����悤�ɂ���.
                cmdKakutei.Enabled = False
            Else
                Me.txtBAKYCD.Enabled = True
                Me.txtBAKYCD.SetFocus
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
    If True = gdDBS.MoveRecords(dbcKeiyakushaMaster, 1) Then '//�f�[�^�� DESC ORDER �������Ă���̂ł���ł悢
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

Private Sub txtBAADJ1_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBAADJ2_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBAADJ3_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBABANK_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBABANK_LostFocus()
    If 0 <= Len(Trim(txtBABANK.Text)) And Len(Trim(txtBABANK.Text)) < 4 Then
        lblBankName.Caption = ""
        Exit Sub
    End If
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Set dyn = gdDBS.SelectBankMaster("DISTINCT DAKJNM", eBankRecordKubun.Bank, Trim(txtBABANK.Text), vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblBankName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))
End Sub

Private Sub txtBAFAXI_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBAFAXJ_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBAFKxx_Change(Index As Integer)
    Call pButtonControl(True)
End Sub

Private Sub txtBAFKxx_DropOpen(Index As Integer, NoDefault As Boolean)
    txtBAFKxx(Index).Calendar.Holidays = gdDBS.Holiday(txtBAFKxx(Index).Year)
End Sub

Private Sub txtBAKJNM_Change()
    If Len(Trim(txtBAKJNM.Text)) = 0 Then
        txtBAKNNM.Text = ""
    End If
    Call pButtonControl(True)
End Sub

Private Sub txtBAKJNM_Furigana(Yomi As String)
    '//���݂̓ǂ݃J�i���ƌ������`�l���������Ȃ�ǂ݃J�i���ƌ������`�l���ɓ]��
    If Trim(txtBAKNNM.Text) = Trim(txtBAKZNM.Text) Then
        txtBAKNNM.Text = txtBAKNNM.Text & Yomi
        txtBAKZNM.Text = txtBAKNNM.Text
    Else
        txtBAKNNM.Text = txtBAKNNM.Text & Yomi
    End If
End Sub

Private Sub txtBAKKRN_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBAKNNM_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBAKSNO_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBAKYCD_KeyDown(KeyCode As Integer, Shift As Integer)
    '// Return �̂Ƃ��̂ݏ�������
    If Not (KeyCode = vbKeyReturn) Then
        Exit Sub
    End If
    
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim msg As String
        
    If "" = Trim(txtBAKYCD.Text) Then
        Exit Sub
    End If
    sql = "SELECT * FROM tbKeiyakushaMaster"
    sql = sql & " WHERE BAITKB = '" & cboABKJNM.ItemData(cboABKJNM.ListIndex) & "'"
    sql = sql & "   AND BAKYCD = '" & txtBAKYCD.Text & "'"
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//    sql = sql & "   AND BAKSCD = '" & txtBAKSCD.Text & "'"
    sql = sql & " ORDER BY BASQNO DESC"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If 0 = dyn.RecordCount Then
        If eShoriKubun.Add <> lblShoriKubun.Caption Then     '���R�[�h�����ŐV�K�ȊO�̎�
            msg = "�Y���f�[�^�͑��݂��܂���."
        End If
    ElseIf eShoriKubun.Add = lblShoriKubun.Caption Then      '���R�[�h�L��ŐV�K�̎�
        msg = "���Ƀf�[�^�����݂��܂�."
    End If
    Set dyn = Nothing
    If msg <> "" Then
        Call MsgBox(msg, vbInformation, mCaption)
'        Call txtBAKYCD.SetFocus
        Exit Sub
    End If
    mIsActivated = False    '//���R�[�h�\�����̃C�x���g���E��Ȃ��悤�Ƀt���O��ݒ�
    dbcKeiyakushaMaster.RecordSource = sql
    Call dbcKeiyakushaMaster.Refresh
    On Error GoTo txtBAKYCD_KeyDownError        '//�r������p�G���[�g���b�v
    If 0& = dbcKeiyakushaMaster.Recordset.RecordCount Then
        '//�V�K�o�^
        dbcKeiyakushaMaster.Recordset.AddNew
        lblBAITKB.Caption = cboABKJNM.ItemData(cboABKJNM.ListIndex)
        lblBAKYCD.Caption = txtBAKYCD.Text
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//        lblBAKSCD.Caption = txtBAKSCD.Text
        lblBAKSCD.Caption = "000"   '// ALL-ZERO �ݒ�
        lblBASQNO.Caption = gdDBS.sysDate("yyyymmdd")
        lblBAKKBN.Caption = 0
        lblBAKZSB.Caption = 1
        '//�_����ԁE�U�����Ԃ̏I������ݒ�
        txtBAKYxx(0).Number = 20000101 '//��U�l��ݒ肵�Ȃ��Ɓu�O�v���Z�b�g����Ȃ��F�s�v�c�H
        txtBAKYxx(0).Number = 0
        txtBAKYxx(1).Number = gdDBS.LastDay(0)
        txtBAFKxx(0).Number = 20000101 '//��U�l��ݒ肵�Ȃ��Ɓu�O�v���Z�b�g����Ȃ��F�s�v�c�H
        txtBAFKxx(0).Number = 0
        txtBAFKxx(1).Number = gdDBS.LastDay(0)
    Else
        If eBankKubun.KinnyuuKikan = dbcKeiyakushaMaster.Recordset.Fields("BAKKBN").Value Then
            '//2007/06/06   ��s���E�x�X���̓ǂݍ��݂������ł���悤�ɕύX
            '//             �Ǎ��ݎ��� Change()=���̕\�� �C�x���g���Ԃ� �x�X�R�[�h�E��s�R�[�h�̏��ɂȂ�x�X�����\������Ȃ����Ƃ�����
            Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Bank, _
                dbcKeiyakushaMaster.Recordset.Fields("BABANK").Value, vDate:=gdDBS.sysDate("YYYYMMDD"))
            lblBankName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))
            Set dyn = Nothing
            Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Shiten, _
                dbcKeiyakushaMaster.Recordset.Fields("BABANK").Value, _
                dbcKeiyakushaMaster.Recordset.Fields("BASITN").Value, vDate:=gdDBS.sysDate("YYYYMMDD"))
            lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"�x�X��_����" �œǂ߂Ȃ�
            Set dyn = Nothing
        End If
        '//�C���E�폜
        Call dbcKeiyakushaMaster.Recordset.MoveFirst
        Call dbcKeiyakushaMaster.Recordset.Edit
'        Call dbcKeiyakushaMaster.UpdateRecord
    End If
    '//�Q�ƂŖ�����΃{�^���̐���J�n
    If False = optShoriKubun(eShoriKubun.Refer).Value Then
        Call pLockedControl(False)
    End If
    spnRireki.Visible = dbcKeiyakushaMaster.Recordset.RecordCount > 1
    '//���̃{�^���͎x�X���N���b�N�������Ɏg����悤�ɂ���.
    cmdKakutei.Enabled = False
    '//�R���g���[���������ԍ��ɂ����������߂ɂ��܂��Ȃ��F���ɕ��@��������Ȃ��H
    If True = optShoriKubun(eShoriKubun.Refer).Value Then
        Call SendKeys("+{TAB}")
    Else
        Call SendKeys("+{TAB}+{TAB}")
    End If
    '//���~�{�^���͎Q�ƈȊO�͂��ł������\�ɁI
    Call pButtonControl(optShoriKubun(eShoriKubun.Delete).Value, True)
    '//���~�{�^���͂��ł������\�ɁI
    If Not optShoriKubun(eShoriKubun.Refer).Value Then
        cmdCancel.Visible = True
        cmdCancel.Enabled = True
    End If
    Exit Sub
txtBAKYCD_KeyDownError:
    Call gdDBS.ErrorCheck(dbcKeiyakushaMaster.Database)    '//�r������p�G���[�g���b�v
End Sub

Private Sub txtBAKYxx_Change(Index As Integer)
    Call pButtonControl(True)
End Sub

Private Sub txtBAKYxx_DropOpen(Index As Integer, NoDefault As Boolean)
    txtBAKYxx(Index).Calendar.Holidays = gdDBS.Holiday(txtBAKYxx(Index).Year)
End Sub

Private Function pUpdateErrorCheck() As Boolean
    '///////////////////////////////
    '//�K�{���͍��ڂƐ������`�F�b�N
    
    Dim str As New StringClass
    Dim obj As Object, msg As String
    '//�ی�ҁE�������͕̂K�{
    If txtBAKJNM.Text = "" Then
        Set obj = txtBAKJNM
        msg = "�_��Җ�(����)�͕K�{���͂ł�."
    ElseIf False = str.CheckLength(txtBAKJNM.Text) Then
        Set obj = txtBAKJNM
        msg = "�_��Җ�(����)�ɔ��p���܂܂�Ă��܂�."
    End If
    '//�ی�ҁE�J�i���͕̂K�{
    If txtBAKNNM.Text = "" Then
        Set obj = txtBAKNNM
        msg = "�_��Җ�(�J�i)�͕K�{���͂ł�."
    ElseIf False = str.CheckLength(txtBAKNNM.Text, vbNarrow) Then
        Set obj = txtBAKNNM
        msg = "�_��Җ�(�J�i)�ɑS�p���܂܂�Ă��܂�."
    ElseIf 0 < InStr(txtBAKNNM.Text, "�") Then
        Set obj = txtBAKNNM
        msg = "�_��Җ�(�J�i)�ɒ������܂܂�Ă��܂�."
    End If
    '//�Z�����̑S�p�`�F�b�N
    If False = str.CheckLength(txtBAADJ1.Text) Then
        Set obj = txtBAADJ1
        msg = "�Z���P(����)�ɔ��p���܂܂�Ă��܂�."
    ElseIf False = str.CheckLength(txtBAADJ2.Text) Then
        Set obj = txtBAADJ2
        msg = "�Z���Q(����)�ɔ��p���܂܂�Ă��܂�."
    ElseIf False = str.CheckLength(txtBAADJ3.Text) Then
        Set obj = txtBAADJ3
        msg = "�Z���R(����)�ɔ��p���܂܂�Ă��܂�."
    End If
    
    If IsNull(txtBAKYxx(1).Number) Then
        Set obj = txtBAKYxx(1)
        msg = "�_����Ԃ̏I�����͕K�{���͂ł�."
    ElseIf txtBAKYxx(0).Text > txtBAKYxx(1).Text Then
        Set obj = txtBAKYxx(0)
        msg = "�_����Ԃ��s���ł�."
    ElseIf IsNull(txtBAFKxx(1).Number) Then
        Set obj = txtBAFKxx(1)
        msg = "�U�����Ԃ̏I�����͕K�{���͂ł�."
    ElseIf txtBAFKxx(0).Text > txtBAFKxx(1).Text Then
        Set obj = txtBAFKxx(0)
        msg = "�U�����Ԃ��s���ł�."
    End If
    
    If lblBAKKBN.Caption = eBankKubun.KinnyuuKikan Then
        If txtBABANK.Text = "" Or lblBankName.Caption = "" Then
            Set obj = txtBABANK
            msg = "���Z�@�ւ͕K�{���͂ł�."
        ElseIf txtBASITN.Text = "" Or lblShitenName.Caption = "" Then
            Set obj = txtBASITN
            msg = "�x�X�͕K�{���͂ł�."
        ElseIf Not (lblBAKZSB.Caption = eBankYokinShubetsu.Futsuu _
                 Or lblBAKZSB.Caption = eBankYokinShubetsu.Touza) Then
            Set obj = optBAKZSB(eBankYokinShubetsu.Futsuu)
            msg = "�a����ʂ͕K�{���͂ł�."
        ElseIf txtBAKZNO.Text = "" Then
            Set obj = txtBAKZNO
            msg = "�����ԍ��͕K�{���͂ł�."
        End If
    ElseIf lblBAKKBN.Caption = eBankKubun.YuubinKyoku Then
        If txtBAYBTK.Text = "" Then
            Set obj = txtBAYBTK
            msg = "�ʒ��L���͕K�{���͂ł�."
        ElseIf txtBAYBTN.Text = "" Then
            Set obj = txtBAYBTN
            msg = "�ʒ��ԍ��͕K�{���͂ł�."
        ElseIf "1" <> Right(txtBAYBTN.Text, 1) Then
'//2006/04/26 �����ԍ��`�F�b�N
            Set obj = txtBAYBTN
            msg = "�ʒ��ԍ��̖������u�P�v�ȊO�ł�."
        End If
    End If
    If txtBAKZNM.Text = "" Then
        Set obj = txtBAKZNM
        msg = "�������`�l��(�J�i)�͕K�{���͂ł�."
    End If
    '//Object ���ݒ肳��Ă��邩�H
    If TypeName(obj) <> "Nothing" Then
        Call MsgBox(msg, vbCritical, mCaption)
        Call obj.SetFocus
        Exit Function
    End If
    
    If lblBASQNO.Caption = gdDBS.sysDate("yyyymmdd") Then
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
    
    sql = "SELECT * FROM tbKeiyakushaMaster"
    sql = sql & " WHERE BAITKB = '" & lblBAITKB.Caption & "'"
    sql = sql & "   AND BAKYCD = '" & lblBAKYCD.Caption & "'"
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//    sql = sql & "   AND BAKSCD = '" & lblBAKSCD.Caption & "'"
    sql = sql & "   AND BASQNO =  " & lblBASQNO.Caption
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If dyn.EOF Then
        Exit Function   '//�V�K�o�^�Ȃ̂Ń`�F�b�N����
    End If
        
    If txtBAKJNM.Text <> gdDBS.Nz(dyn.Fields("BAKJNM")) _
    Or txtBAKNNM.Text <> gdDBS.Nz(dyn.Fields("BAKNNM")) Then
        AddRireki = "�_���"
    ElseIf lblBAKKBN.Caption <> gdDBS.Nz(dyn.Fields("BAKKBN")) Then
        AddRireki = "�U�֌���"
    ElseIf lblBAKKBN.Caption = eBankKubun.KinnyuuKikan Then
        '//���Z�@�֏�񂪈Ⴆ�Η������ǉ�
        If txtBABANK.Text <> gdDBS.Nz(dyn.Fields("BABANK")) _
         Or txtBASITN.Text <> gdDBS.Nz(dyn.Fields("BASITN")) _
         Or lblBAKZSB.Caption <> gdDBS.Nz(dyn.Fields("BAKZSB")) _
         Or txtBAKZNO.Text <> gdDBS.Nz(dyn.Fields("BAKZNO")) Then
            AddRireki = "���ԋ@��"
        End If
    ElseIf lblBAKKBN.Caption = eBankKubun.YuubinKyoku Then
        '//�X�֋Ǐ�񂪈Ⴆ�Η������ǉ�
        If txtBAYBTK.Text <> gdDBS.Nz(dyn.Fields("BAYBTK")) _
         Or txtBAYBTN.Text <> gdDBS.Nz(dyn.Fields("BAYBTN")) Then
            AddRireki = "�X�֋�"
        End If
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
        .lblFurikomi.Caption = "�U���J�n��"
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
    sql = " DELETE tbKeiyakushaMaster"
    sql = sql & " WHERE BAITKB = '" & lblBAITKB.Caption & "'"
    sql = sql & "   AND BAKYCD = '" & lblBAKYCD.Caption & "'"
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//    sql = sql & "   AND BAKSCD = '" & lblBAKSCD.Caption & "'"
    sql = sql & "   AND BASQNO = -1"
    Call gdDBS.Database.ExecuteSQL(sql)
    
    '////////////////////////////////////////////////
    '//�e�[�u����`���ύX���ꂽ�ꍇ���ӂ��邱�ƁI�I
    Dim FixedCol As String
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//    FixedCol = "BAITKB,BAKYCD,BAKSCD,BAKSNM,BAKSNO,BAKJNM,BAKNNM," &
    FixedCol = "BAITKB,BAKYCD,BAKSNM,BAKSNO,BAKJNM,BAKNNM," & _
               "BAZPC1,BAZPC2,BAADJ1,BAADJ2,BAADJ3,BATELE,BAKKRN," & _
               "BATELJ,BAFAXI,BAKKBN,BABANK,BAFAXJ,BASITN,BAKZSB," & _
               "BAKZNO,BAKZNM,BAYBTK,BAYBTN,BAKYST,BAFKST,BAKYFG," & _
               "BASCNT,BAUSID,BAADDT"
    '���݂̍X�V�O�f�[�^�ޔ�
    sql = "INSERT INTO tbKeiyakushaMaster("
'//2012/08/10 �����敪(??KSCD) NOT NULL ����ŃG���[�ɂȂ�̂ŕ����FNULL �Ȃ�=>000
    sql = sql & "BAKSCD,BASQNO,BAKYED,BAFKED,BAUPDT,"
    sql = sql & FixedCol
    sql = sql & ") SELECT "
    sql = sql & "NVL(BAKSCD,'000'),-1,"
    '//���͂��ꂽ���̑O��������ݒ�
    sql = sql & "TO_CHAR(TO_DATE(" & KeiyakuEnd & ",'YYYYMMDD')-1,'YYYYMMDD'),"
    sql = sql & "TO_CHAR(TO_DATE(" & FurikaeEnd & ",'YYYYMMDD')-1,'YYYYMMDD'),"
    sql = sql & " SYSDATE,"
    sql = sql & FixedCol
    sql = sql & " FROM tbKeiyakushaMaster"
    sql = sql & " WHERE BAITKB = '" & lblBAITKB.Caption & "'"
    sql = sql & "   AND BAKYCD = '" & lblBAKYCD.Caption & "'"
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//    sql = sql & "   AND BAKSCD = '" & lblBAKSCD.Caption & "'"
    sql = sql & "   AND BASQNO =  " & lblBASQNO.Caption
    Call gdDBS.Database.ExecuteSQL(sql)
    
    txtBAKYxx(0).Number = KeiyakuEnd
    txtBAFKxx(0).Number = FurikaeEnd
    
    '//��ʂ̓��e���X�V:cmdUpdate()�̈ꕔ�֐������s
    Call pUpdateRecord
    
    On Error GoTo pRirekiAddNewError
    '//��ʂ̃f�[�^�̂r�d�p��{���ɂ���
    sql = "UPDATE tbKeiyakushaMaster SET "
    sql = sql & "BASQNO = TO_CHAR(SYSDATE,'YYYYMMDD'),"
    sql = sql & "BAUSID = '" & gdDBS.LoginUserName & "',"
    sql = sql & "BAUPDT = SYSDATE"
    sql = sql & " WHERE BAITKB = '" & lblBAITKB.Caption & "'"
    sql = sql & "   AND BAKYCD = '" & lblBAKYCD.Caption & "'"
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//    sql = sql & "   AND BAKSCD = '" & lblBAKSCD.Caption & "'"
    sql = sql & "   AND BASQNO =  " & lblBASQNO.Caption
    Call gdDBS.Database.ExecuteSQL(sql)
    '//�ޔ������f�[�^�̂r�d�p��ύX�O�ɂ���
    sql = "UPDATE tbKeiyakushaMaster SET "
    sql = sql & "BASQNO = " & lblBASQNO.Caption & ","
    sql = sql & "BAUSID = '" & gdDBS.LoginUserName & "',"
    sql = sql & "BAUPDT = SYSDATE"
    sql = sql & " WHERE BAITKB = '" & lblBAITKB.Caption & "'"
    sql = sql & "   AND BAKYCD = '" & lblBAKYCD.Caption & "'"
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//    sql = sql & "   AND BAKSCD = '" & lblBAKSCD.Caption & "'"
    sql = sql & "   AND BASQNO = -1"
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

Private Sub txtBAKZNM_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBAKZNO_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBASITN_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBASITN_LostFocus()
    If 0 <= Len(Trim(txtBASITN.Text)) And Len(Trim(txtBASITN.Text)) < 3 Then
        lblShitenName.Caption = ""
        Exit Sub
    End If
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Shiten, Trim(txtBABANK.Text), Trim(txtBASITN.Text), vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"�x�X��_����" �œǂ߂Ȃ�
End Sub

Private Sub txtBATELE_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBATELJ_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBAYBTK_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBAYBTK_LostFocus()
'//2006/04/26 �O�[�����ߍ���
    If "" <> txtBAYBTK.Text Then
        txtBAYBTK.Text = Format(Val(txtBAYBTK.Text), "000")
    End If
End Sub

Private Sub txtBAYBTN_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBAYBTN_LostFocus()
    '//2006/04/26 �O�[�����ߍ���
    If "" <> txtBAYBTN.Text Then
        If "1" <> Right(txtBAYBTN.Text, 1) Then
            Call MsgBox("�������u�P�v�ȊO�ł�.(" & lblTsuchoBango.Caption & ")", vbCritical, mCaption)
        Else
            txtBAYBTN.Text = Format(Val(txtBAYBTN.Text), "00000000")
        End If
    End If
End Sub

Private Sub txtBAZPC1_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtBAZPC2_Change()
    Call pButtonControl(True)
End Sub
