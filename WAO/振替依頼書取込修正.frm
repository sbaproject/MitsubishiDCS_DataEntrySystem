VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmFurikaeReqImportEdit 
   Caption         =   "�U�ֈ˗���(�捞)�C��"
   ClientHeight    =   7710
   ClientLeft      =   4455
   ClientTop       =   3060
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
   ScaleHeight     =   7710
   ScaleWidth      =   10125
   Begin VB.Frame fraSysDate 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      Height          =   375
      Left            =   8640
      TabIndex        =   81
      Top             =   -60
      Width           =   1155
      Begin VB.Label lblSysDate 
         Caption         =   "Label1"
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
         TabIndex        =   82
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "�߂�(&B)"
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
      Left            =   8220
      TabIndex        =   78
      Top             =   6945
      Width           =   1335
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
      Left            =   4140
      TabIndex        =   76
      Top             =   6945
      Width           =   1335
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
      Left            =   5640
      TabIndex        =   77
      Top             =   6945
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "�O�̃f�[�^(&P)"
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
      Left            =   720
      TabIndex        =   74
      Top             =   6945
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "���̃f�[�^(&N)"
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
      Left            =   2220
      TabIndex        =   75
      Top             =   6945
      Width           =   1335
   End
   Begin VB.ComboBox cboCIOKFG 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      ItemData        =   "�U�ֈ˗����捞�C��.frx":0000
      Left            =   1800
      List            =   "�U�ֈ˗����捞�C��.frx":000D
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   59
      Top             =   5010
      Width           =   2835
   End
   Begin VB.CheckBox chkCIMUPD 
      Caption         =   "�}�X�^���f���Ȃ�"
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
      Left            =   2640
      TabIndex        =   57
      Top             =   4650
      Width           =   1935
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
      ItemData        =   "�U�ֈ˗����捞�C��.frx":003D
      Left            =   1800
      List            =   "�U�ֈ˗����捞�C��.frx":004A
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      TabStop         =   0   'False
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
      TabIndex        =   10
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
         Left            =   420
         TabIndex        =   22
         Top             =   1260
         Width           =   4035
         Begin imText6Ctl.imText txtCiYBTK 
            DataField       =   "CiYBTK"
            DataSource      =   "dbcImportEdit"
            Height          =   285
            Left            =   1860
            TabIndex        =   23
            Top             =   480
            Width           =   375
            _Version        =   65537
            _ExtentX        =   661
            _ExtentY        =   503
            Caption         =   "�U�ֈ˗����捞�C��.frx":0068
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�U�ֈ˗����捞�C��.frx":00D4
            Key             =   "�U�ֈ˗����捞�C��.frx":00F2
            MouseIcon       =   "�U�ֈ˗����捞�C��.frx":0136
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
         Begin imText6Ctl.imText txtCiYBTN 
            DataField       =   "CiYBTN"
            DataSource      =   "dbcImportEdit"
            Height          =   285
            Left            =   1860
            TabIndex        =   24
            Top             =   960
            Width           =   855
            _Version        =   65537
            _ExtentX        =   1508
            _ExtentY        =   503
            Caption         =   "�U�ֈ˗����捞�C��.frx":0152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�U�ֈ˗����捞�C��.frx":01BE
            Key             =   "�U�ֈ˗����捞�C��.frx":01DC
            MouseIcon       =   "�U�ֈ˗����捞�C��.frx":0220
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
            TabIndex        =   40
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
            TabIndex        =   39
            Top             =   480
            Width           =   1275
         End
      End
      Begin VB.OptionButton optCiKKBN 
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optCiKKBN 
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
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   1395
      End
      Begin imText6Ctl.imText txtCiKZNM 
         DataField       =   "CiKZNM"
         DataSource      =   "dbcImportEdit"
         Height          =   285
         Left            =   420
         TabIndex        =   25
         Top             =   2580
         Width           =   3735
         _Version        =   65537
         _ExtentX        =   6588
         _ExtentY        =   503
         Caption         =   "�U�ֈ˗����捞�C��.frx":023C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "�U�ֈ˗����捞�C��.frx":02A8
         Key             =   "�U�ֈ˗����捞�C��.frx":02C6
         MouseIcon       =   "�U�ֈ˗����捞�C��.frx":030A
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
         TabIndex        =   15
         Top             =   420
         Width           =   3855
         Begin imText6Ctl.imText txtCiKZNO 
            DataField       =   "CiKZNO"
            DataSource      =   "dbcImportEdit"
            Height          =   285
            Left            =   1140
            TabIndex        =   21
            Top             =   1380
            Width           =   795
            _Version        =   65537
            _ExtentX        =   1402
            _ExtentY        =   503
            Caption         =   "�U�ֈ˗����捞�C��.frx":0326
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�U�ֈ˗����捞�C��.frx":0392
            Key             =   "�U�ֈ˗����捞�C��.frx":03B0
            MouseIcon       =   "�U�ֈ˗����捞�C��.frx":03F4
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
         Begin imText6Ctl.imText txtCiSITN 
            DataField       =   "CiSITN"
            DataSource      =   "dbcImportEdit"
            Height          =   285
            Left            =   1200
            TabIndex        =   17
            Top             =   660
            Width           =   375
            _Version        =   65537
            _ExtentX        =   661
            _ExtentY        =   503
            Caption         =   "�U�ֈ˗����捞�C��.frx":0410
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�U�ֈ˗����捞�C��.frx":047C
            Key             =   "�U�ֈ˗����捞�C��.frx":049A
            MouseIcon       =   "�U�ֈ˗����捞�C��.frx":04DE
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
         Begin imText6Ctl.imText txtCiBANK 
            DataField       =   "CiBANK"
            DataSource      =   "dbcImportEdit"
            Height          =   285
            Left            =   1200
            TabIndex        =   16
            Top             =   300
            Width           =   495
            _Version        =   65537
            _ExtentX        =   873
            _ExtentY        =   503
            Caption         =   "�U�ֈ˗����捞�C��.frx":04FA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "�U�ֈ˗����捞�C��.frx":0566
            Key             =   "�U�ֈ˗����捞�C��.frx":0584
            MouseIcon       =   "�U�ֈ˗����捞�C��.frx":05C8
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
            TabIndex        =   41
            Top             =   900
            Width           =   2535
            Begin VB.OptionButton optCiKZSB 
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
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   180
               Width           =   675
            End
            Begin VB.OptionButton optCiKZSB 
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
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   180
               Width           =   675
            End
            Begin VB.OptionButton optCiKZSB 
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
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   480
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label lblCiKZSB 
               BackColor       =   &H000000FF&
               Caption         =   "�������"
               DataField       =   "CiKZSB"
               DataSource      =   "dbcImportEdit"
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
               TabIndex        =   42
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
            Top             =   660
            Width           =   1935
         End
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
         TabIndex        =   55
         Top             =   2340
         Width           =   1395
      End
      Begin VB.Label lblCiKKBN 
         BackColor       =   &H000000FF&
         Caption         =   "���Z�@�֎��"
         DataField       =   "CiKKBN"
         DataSource      =   "dbcImportEdit"
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
         TabIndex        =   49
         Top             =   180
         Width           =   1095
      End
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
      TabIndex        =   26
      Top             =   3360
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
         ItemData        =   "�U�ֈ˗����捞�C��.frx":05E4
         Left            =   1500
         List            =   "�U�ֈ˗����捞�C��.frx":0609
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   27
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
         ItemData        =   "�U�ֈ˗����捞�C��.frx":064B
         Left            =   3900
         List            =   "�U�ֈ˗����捞�C��.frx":0670
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   29
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
         TabIndex        =   31
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
         DatabaseName    =   "srv0905"
         Connect         =   "wao/wao"
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
         DatabaseName    =   "srv0905"
         Connect         =   "wao/wao"
         RecordSource    =   ""
      End
      Begin MSDBCtls.DBList dblBankList 
         Bindings        =   "�U�ֈ˗����捞�C��.frx":06B2
         Height          =   2040
         Left            =   120
         TabIndex        =   28
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
         Bindings        =   "�U�ֈ˗����捞�C��.frx":06C8
         Height          =   2040
         Left            =   2400
         TabIndex        =   30
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
         TabIndex        =   51
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
         TabIndex        =   50
         Top             =   240
         Width           =   1395
      End
   End
   Begin ORADCLibCtl.ORADC dbcImportEdit 
      Height          =   315
      Left            =   4740
      Top             =   7995
      Visible         =   0   'False
      Width           =   2235
      _Version        =   65536
      _ExtentX        =   3942
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
      DatabaseName    =   "srv0905"
      Connect         =   "wao2/wao2"
      RecordSource    =   "SELECT * FROM tcHogoshaImport WHERE rowid is null"
   End
   Begin imDate6Ctl.imDate txtCiFKxx 
      DataField       =   "CIFKST"
      DataSource      =   "dbcImportEdit"
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   6
      Top             =   3255
      Width           =   795
      _Version        =   65537
      _ExtentX        =   1402
      _ExtentY        =   556
      Calendar        =   "�U�ֈ˗����捞�C��.frx":06E0
      Caption         =   "�U�ֈ˗����捞�C��.frx":0860
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�U�ֈ˗����捞�C��.frx":08CC
      Keys            =   "�U�ֈ˗����捞�C��.frx":08EA
      MouseIcon       =   "�U�ֈ˗����捞�C��.frx":0948
      Spin            =   "�U�ֈ˗����捞�C��.frx":0964
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
   Begin imText6Ctl.imText txtCiKJNM 
      DataField       =   "CiKJNM"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1995
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "�U�ֈ˗����捞�C��.frx":098C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�U�ֈ˗����捞�C��.frx":09F8
      Key             =   "�U�ֈ˗����捞�C��.frx":0A16
      MouseIcon       =   "�U�ֈ˗����捞�C��.frx":0A5A
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
   Begin imText6Ctl.imText txtCiKYCD 
      DataField       =   "CIKYCD"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   1320
      Width           =   795
      _Version        =   65537
      _ExtentX        =   1402
      _ExtentY        =   503
      Caption         =   "�U�ֈ˗����捞�C��.frx":0A76
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�U�ֈ˗����捞�C��.frx":0AE2
      Key             =   "�U�ֈ˗����捞�C��.frx":0B00
      MouseIcon       =   "�U�ֈ˗����捞�C��.frx":0B44
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
   Begin imText6Ctl.imText txtCiHGCD 
      DataField       =   "CIHGCD"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   1635
      Width           =   855
      _Version        =   65537
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "�U�ֈ˗����捞�C��.frx":0B60
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�U�ֈ˗����捞�C��.frx":0BCC
      Key             =   "�U�ֈ˗����捞�C��.frx":0BEA
      MouseIcon       =   "�U�ֈ˗����捞�C��.frx":0C2E
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
      Left            =   4740
      Top             =   7575
      Visible         =   0   'False
      Width           =   2235
      _Version        =   65536
      _ExtentX        =   3942
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
      DatabaseName    =   "srv0905"
      Connect         =   "wao/wao"
      RecordSource    =   "SELECT * FROM taItakushaMaster"
   End
   Begin imText6Ctl.imText txtCiSTNM 
      DataField       =   "CiSTNM"
      DataSource      =   "dbcImportEdit"
      Height          =   465
      Left            =   1800
      TabIndex        =   5
      Top             =   2715
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   820
      Caption         =   "�U�ֈ˗����捞�C��.frx":0C4A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�U�ֈ˗����捞�C��.frx":0CB6
      Key             =   "�U�ֈ˗����捞�C��.frx":0CD4
      MouseIcon       =   "�U�ֈ˗����捞�C��.frx":0D18
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
   Begin imText6Ctl.imText txtCiKNNM 
      DataField       =   "CiKNNM"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   2355
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "�U�ֈ˗����捞�C��.frx":0D34
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�U�ֈ˗����捞�C��.frx":0DA0
      Key             =   "�U�ֈ˗����捞�C��.frx":0DBE
      MouseIcon       =   "�U�ֈ˗����捞�C��.frx":0E02
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
   Begin imDate6Ctl.imDate txtCiFKxx 
      DataSource      =   "dbcImportEdit"
      Height          =   315
      Index           =   1
      Left            =   3120
      TabIndex        =   7
      Top             =   3255
      Visible         =   0   'False
      Width           =   795
      _Version        =   65537
      _ExtentX        =   1402
      _ExtentY        =   556
      Calendar        =   "�U�ֈ˗����捞�C��.frx":0E1E
      Caption         =   "�U�ֈ˗����捞�C��.frx":0F9E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�U�ֈ˗����捞�C��.frx":100A
      Keys            =   "�U�ֈ˗����捞�C��.frx":1028
      MouseIcon       =   "�U�ֈ˗����捞�C��.frx":1086
      Spin            =   "�U�ֈ˗����捞�C��.frx":10A2
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
   Begin imText6Ctl.imText txtCIWMSG 
      Height          =   1455
      Left            =   780
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   5370
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   2566
      Caption         =   "�U�ֈ˗����捞�C��.frx":10CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�U�ֈ˗����捞�C��.frx":1126
      Key             =   "�U�ֈ˗����捞�C��.frx":1144
      BackColor       =   -2147483633
      EditMode        =   1
      ForeColor       =   16711935
      ReadOnly        =   1
      ShowContextMenu =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   0
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   -1
      ScrollBars      =   3
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   "�x�����b�Z�[�W�������s�ɕ\�������B"
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   3
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   1
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin imText6Ctl.imText txtCiBKNM 
      DataField       =   "CiBKNM"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1815
      TabIndex        =   8
      Top             =   3615
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "�U�ֈ˗����捞�C��.frx":1188
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�U�ֈ˗����捞�C��.frx":11F4
      Key             =   "�U�ֈ˗����捞�C��.frx":1212
      MouseIcon       =   "�U�ֈ˗����捞�C��.frx":1256
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
      Text            =   "���Z�@�֖��D�D�D�D�D�D�D�D�D��"
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
   Begin imText6Ctl.imText txtCiSINM 
      DataField       =   "CiSINM"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1815
      TabIndex        =   9
      Top             =   3975
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "�U�ֈ˗����捞�C��.frx":1272
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�U�ֈ˗����捞�C��.frx":12DE
      Key             =   "�U�ֈ˗����捞�C��.frx":12FC
      MouseIcon       =   "�U�ֈ˗����捞�C��.frx":1340
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
      Text            =   "�x�X���D�D�D�D�D�D�D�D�D�D�D��"
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
   Begin VB.Frame fraCiINSD 
      Enabled         =   0   'False
      Height          =   465
      Left            =   1800
      TabIndex        =   83
      Top             =   4200
      Width           =   2040
      Begin VB.OptionButton optCiINSD 
         Caption         =   "�u����"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   85
         Top             =   150
         Width           =   900
      End
      Begin VB.OptionButton optCiINSD 
         Caption         =   "�ǉ�"
         Height          =   240
         Index           =   1
         Left            =   1200
         TabIndex        =   84
         Top             =   150
         Width           =   690
      End
      Begin VB.Label lblCiINSD 
         BackColor       =   &H000000FF&
         Caption         =   "�X�V���@"
         DataField       =   "CiINSD"
         DataSource      =   "dbcImportEdit"
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
         Left            =   0
         TabIndex        =   87
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Label lblUpdMode 
      Alignment       =   1  '�E����
      Caption         =   "�X�V���@"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   300
      TabIndex        =   86
      Top             =   4350
      Width           =   1395
   End
   Begin VB.Label Label8 
      Alignment       =   1  '�E����
      Caption         =   "�x�X��(�捞)"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   300
      TabIndex        =   80
      Top             =   4020
      Width           =   1395
   End
   Begin VB.Label Label9 
      Alignment       =   1  '�E����
      Caption         =   "���Z�@�֖�(�捞)"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   300
      TabIndex        =   79
      Top             =   3645
      Width           =   1395
   End
   Begin VB.Label lblCIUPDT 
      BackColor       =   &H000000FF&
      Caption         =   "�X�V��"
      DataField       =   "CIUPDT"
      DataSource      =   "dbcImportEdit"
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
      Left            =   3300
      TabIndex        =   73
      Top             =   7995
      Width           =   915
   End
   Begin VB.Label lblCIUSID 
      BackColor       =   &H000000FF&
      Caption         =   "�X�V��"
      DataField       =   "CIUSID"
      DataSource      =   "dbcImportEdit"
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
      Left            =   2280
      TabIndex        =   72
      Top             =   7995
      Width           =   795
   End
   Begin VB.Label lblCIWMSG 
      BackColor       =   &H000000FF&
      Caption         =   "�x�����b�Z�[�W�������s�ɕ\�������B"
      DataField       =   "CIWMSG"
      DataSource      =   "dbcImportEdit"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   660
      TabIndex        =   71
      Top             =   7395
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Label lblCIERROR 
      BackColor       =   &H000000FF&
      Caption         =   "�ύX��-F"
      DataField       =   "CIERROR"
      DataSource      =   "dbcImportEdit"
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
      TabIndex        =   70
      Top             =   7635
      Width           =   795
   End
   Begin VB.Label lblCIERSR 
      BackColor       =   &H000000FF&
      Caption         =   "�ύX�O-F"
      DataField       =   "CIERSR"
      DataSource      =   "dbcImportEdit"
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
      Left            =   720
      TabIndex        =   69
      Top             =   7635
      Width           =   795
   End
   Begin VB.Label lblCIOKFG 
      BackColor       =   &H000000FF&
      Caption         =   "���f�n�j"
      DataField       =   "CIOKFG"
      DataSource      =   "dbcImportEdit"
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
      Left            =   2400
      TabIndex        =   68
      Top             =   7635
      Width           =   855
   End
   Begin VB.Label lblCIMUPD 
      BackColor       =   &H000000FF&
      Caption         =   "���f���Ȃ�"
      DataField       =   "CIMUPD"
      DataSource      =   "dbcImportEdit"
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
      Left            =   3360
      TabIndex        =   67
      Top             =   7635
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�捞����-SEQ"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   300
      TabIndex        =   66
      Top             =   525
      Width           =   1395
   End
   Begin VB.Label lblCISEQN 
      Alignment       =   1  '�E����
      BorderStyle     =   1  '����
      Caption         =   "�捞SEQ"
      DataField       =   "CISEQN"
      DataSource      =   "dbcImportEdit"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3855
      TabIndex        =   65
      Top             =   480
      Width           =   795
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label71 
      Alignment       =   1  '�E����
      Caption         =   "�|"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3630
      TabIndex        =   64
      Top             =   540
      Width           =   180
   End
   Begin VB.Label lblCIINDT 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      Caption         =   "2006/03/01 23:59:59"
      DataField       =   "CIINDT"
      DataSource      =   "dbcImportEdit"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1815
      TabIndex        =   63
      Top             =   480
      Width           =   1755
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgCIWMSG 
      Appearance      =   0  '�ׯ�
      Height          =   480
      Left            =   240
      Picture         =   "�U�ֈ˗����捞�C��.frx":135C
      Top             =   5370
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label10x 
      Alignment       =   1  '�E����
      Caption         =   "�}�X�^���f���@"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   285
      TabIndex        =   62
      Top             =   5070
      Width           =   1395
   End
   Begin VB.Label lblERRMSG 
      Alignment       =   2  '��������
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "�ُ�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   61
      Top             =   4650
      Width           =   555
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label11 
      Alignment       =   1  '�E����
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   285
      TabIndex        =   60
      Top             =   4695
      Width           =   1395
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
      TabIndex        =   56
      Top             =   2355
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
      TabIndex        =   54
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label lblCiITKB 
      BackColor       =   &H000000FF&
      Caption         =   "�ϑ��ҋ敪"
      DataField       =   "CiITKB"
      DataSource      =   "dbcImportEdit"
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
      Left            =   3660
      TabIndex        =   53
      Top             =   840
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
      TabIndex        =   52
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label lblCiSQNO 
      BackColor       =   &H000000FF&
      Caption         =   "�ی�҂r�d�p"
      DataField       =   "CiSQNO"
      DataSource      =   "dbcImportEdit"
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
      Top             =   1725
      Width           =   975
   End
   Begin VB.Label lblCiKYCD 
      BackColor       =   &H000000FF&
      Caption         =   "�I�[�i�[�ԍ�"
      DataField       =   "CiKYCD"
      DataSource      =   "dbcImportEdit"
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
      Left            =   3660
      TabIndex        =   38
      Top             =   1140
      Width           =   1155
   End
   Begin VB.Label lblCiHGCD 
      BackColor       =   &H000000FF&
      Caption         =   "�ی�Ҕԍ�"
      DataField       =   "CiHGCD"
      DataSource      =   "dbcImportEdit"
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
      Left            =   2850
      TabIndex        =   13
      Top             =   1725
      Width           =   975
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
      TabIndex        =   37
      Top             =   1380
      Width           =   2235
   End
   Begin VB.Label Label10 
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
      Left            =   2820
      TabIndex        =   36
      Top             =   3315
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
      TabIndex        =   35
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label lblHogoshaCode 
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
      Left            =   480
      TabIndex        =   34
      Top             =   1695
      Width           =   1155
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
      TabIndex        =   33
      Top             =   2040
      Width           =   1395
   End
   Begin VB.Label Label18 
      Alignment       =   1  '�E����
      Caption         =   "�U�֊J�n�N��"
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
      TabIndex        =   32
      Top             =   3315
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
Attribute VB_Name = "frmFurikaeReqImportEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm As New FormClass
Private mCaption As String
Private mBankChange As Boolean  '//2006/08/22 ???_Change �C�x���g����s=>�x�X�ɋ�������

Private mErrMsgOn As Boolean
Private mCheckUpdate As Boolean
Private mRimp As New FurikaeReqImpClass
Private mUpdateOK As Boolean
Private mIsActivated As Boolean

'//2007/06/07 �X�V�E���~�{�^�������S�P�ƂɃR���g���[��
Private Sub pButtonControl(ByVal vMode As Boolean, Optional vExec As Boolean = False)
    If True = mIsActivated Or True = vExec Then
        cmdUpdate.Visible = vMode
        cmdCancel.Visible = vMode
        cmdUpdate.Enabled = vMode
        cmdCancel.Enabled = vMode
        mIsActivated = True
    End If
End Sub

Private Sub pLockedControl(blMode As Boolean)
    Dim used As Boolean
    used = fraCiINSD.Enabled
    
    Call mForm.LockedControl(False) '//��Ƀf�[�^�͕ҏW�\�ɂ��Ă���
'    cmdUpdate.Enabled = blMode
'    cmdCancel.Enabled = blMode
    'mForm.LockedControl() �Ōx���\�����ԐF�ׁ̈A������I
    lblERRMSG.Visible = True
    '//2007/06/07 �������`�l�͏�ɓ��͂��Ȃ��F�ی�Җ�(�J�i)���R�s�[����l�Ɏd�l�ύX
    txtCiKZNM.Enabled = False
    lblKouzaName.Enabled = False
    cmdKakutei.Enabled = Not blMode
    
    fraCiINSD.Enabled = used
End Sub

Private Sub cboBankYomi_Click()
    Call gdDBS.BankDbListRefresh(dbcBank, cboBankYomi, dblBankList, eBankRecordKubun.Bank)
    dbcShiten.RecordSource = ""
    dbcShiten.Refresh
    dblShitenList.ListField = ""
    dblShitenList.Refresh
    cmdKakutei.Enabled = False
End Sub

Private Sub cboABKJNM_Click()
    If 0 <= cboABKJNM.ListIndex Then
        lblCiITKB.Caption = cboABKJNM.ItemData(cboABKJNM.ListIndex)
        '//�L�[�����������ɍX�V�\�����f
'        cmdUpdate.Enabled = mCheckUpdate    '//�X�V�{�^���̐���F�f�[�^�\�����ɃC�x���g���������Ă��\�Ȃ悤�ɁI
'        cmdCancel.Enabled = cmdUpdate.Enabled
    End If
    Call pButtonControl(True)
End Sub

Private Sub cboCIOKFG_Click()
    '//�C���O�̃G���[�ɂ��I����e�𐧌䂷��
    Select Case lblCIERSR.Caption
    Case mRimp.errEditData    '//���肦�Ȃ�
    Case mRimp.errInvalid, mRimp.errImport
        If cboCIOKFG.ItemData(cboCIOKFG.ListIndex) <> mRimp.updInvalid Then
            '//��؂̑I��s�\�I�I�I
            Call MsgBox("�u�捞�v���́u�ُ�v�f�[�^�ׁ̈A�I���ł��܂���B" & vbCrLf & "�`�F�b�N���������s���ĉ������B", vbCritical + vbOKOnly, mCaption)
            '//cboCIOKFG.ListIndex = mRimp.updInvalid + 2     '// -2 �` 2
            '//���ɖ߂�
            cboCIOKFG.ListIndex = Val(lblCIOKFG.Caption) + 2    '// -2 �` 2
            Exit Sub
        End If
    Case mRimp.errNormal
        '//���ł��n�j
        '//2014/06/11 ����ԂŖ����̂ɉ�������I�������ꍇ
        If False = checkKaiyaku() Then
            If cboCIOKFG.ItemData(cboCIOKFG.ListIndex) = mRimp.updResetCancel Then
                '//�������͊֌W�Ȃ�
                Call MsgBox("����Ԃł͂���܂���B", vbInformation + vbOKOnly, mCaption)
                '//���ɖ߂�
                cboCIOKFG.ListIndex = Val(lblCIOKFG.Caption) + 2    '// -2 �` 2
            End If
        End If
    Case mRimp.errWarning
        If cboCIOKFG.ItemData(cboCIOKFG.ListIndex) = mRimp.updNormal Then
            '//�ă`�F�b�N���Ɍx���ɖ߂�̂őI���̈Ӗ�������
            Call MsgBox("�u�x���v�f�[�^�𔽉f����ɂ�" & vbCrLf & "�u" & mRimp.mUpdateMessage(mRimp.updWarnUpd) & "�v��I�����Ă��������B", vbInformation + vbOKOnly, mCaption)
            '//���ɖ߂�
            cboCIOKFG.ListIndex = Val(lblCIOKFG.Caption) + 2    '// -2 �` 2
            Exit Sub
        Else
            If checkKaiyaku() Then
            'If InStr(lblCIWMSG.Caption, "�����") Then
                If cboCIOKFG.ItemData(cboCIOKFG.ListIndex) = mRimp.updWarnUpd Then
                    '//���������Ȃ��ėǂ���
                    If vbOK <> MsgBox("����Ԃ͉�������܂���B" & vbCrLf & "��낵���ł����H", vbInformation + vbOKCancel, mCaption) Then
                        Exit Sub
                    End If
                End If
            ElseIf cboCIOKFG.ItemData(cboCIOKFG.ListIndex) = mRimp.updResetCancel Then
                '//�������͊֌W�Ȃ�
                Call MsgBox("����Ԃł͂���܂���B", vbInformation + vbOKOnly, mCaption)
                '//���ɖ߂�
                cboCIOKFG.ListIndex = Val(lblCIOKFG.Caption) + 2    '// -2 �` 2
            End If
        End If
    Case Else                   '//���肦�Ȃ�
    End Select
    lblCIOKFG.Caption = cboCIOKFG.ItemData(cboCIOKFG.ListIndex)
    '//2014/06/09 �R���{�{�b�N�X�ύX���Ƀ{�^�����g�p�\��
    Call pButtonControl(True)
    '//�L�[�����������ɍX�V�\�����f
'    cmdUpdate.Enabled = mCheckUpdate    '//�X�V�{�^���̐���F�f�[�^�\�����ɃC�x���g���������Ă��\�Ȃ悤�ɁI
'    cmdCancel.Enabled = cmdUpdate.Enabled
    'Call SendKeys("{TAB}")  '//���ʂ𐳂������������̂Ńt�H�[�J�X�ړ�
End Sub

Private Sub cboShitenYomi_Click()
    If dblBankList.Text = "" Then
        Exit Sub
    End If
    Call gdDBS.BankDbListRefresh(dbcShiten, cboShitenYomi, dblShitenList, eBankRecordKubun.Shiten, Left(dblBankList.Text, 4))
    cmdKakutei.Enabled = False
End Sub

Private Sub chkCIMUPD_Click()
    lblCIMUPD.Caption = Abs(Val(chkCIMUPD.Value))
    Call pLockedControl(True)
    Call pButtonControl(True)
    Call sttausCIERROR  '//2014/05/19 �f�[�^���C�x���g���ɂ��낢�딭������̂ł����ɓ���
End Sub

Private Sub cmdCancel_Click()
    Call dbcImportEdit.UpdateControls
    'Call cboABKJNM.SetFocus
    Call pLockedControl(False)
    Call lblCIERROR_Change
    Call pButtonControl(False)
    Call sttausCIERROR  '//2014/05/19 �f�[�^���C�x���g���ɂ��낢�딭������̂ł����ɓ���
End Sub

Private Function pCheckEditData() As Boolean
    Dim obj As Object, Edit As Boolean
    For Each obj In Me.Controls
        If TypeOf obj Is imText _
        Or TypeOf obj Is imNumber _
        Or TypeOf obj Is imDate _
        Or TypeOf obj Is Label Then
            '//�R���g���[���� DataChanged �v���p�e�B���������čX�V��K�v�Ƃ��邩���f
            If "" <> obj.DataField And True = obj.DataChanged Then
                pCheckEditData = True
                Exit Function
            End If
        End If
    Next obj
End Function

Private Sub cmdUpdate_Click()
    If False = pCheckEditData Then
        Call pLockedControl(False)
        Exit Sub
    End If
    '//���͓��e�`�F�b�N�Ŏ���߂����̂ŏI��
    mUpdateOK = pUpdateErrorCheck
    If False = mUpdateOK Then
        Exit Sub
    End If
    mUpdateOK = True
    lblCIERROR.Caption = mRimp.errEditData    '//�ҏW��͕K���G���[�t���O�𗧂Ă�F�`�F�b�N������K������
    lblCIUSID.Caption = gdDBS.LoginUserName
    lblCIUPDT.Caption = gdDBS.sysDate
    '//���C���� SpreadSheet �ɓ��e�𔽉f����FUpdate��ł� DataChanged() ���ω����Ă��܂��̂�...�B
    Call frmFurikaeReqImport.gEditToSpreadSheet(0)
    '//��ʂ̓��e���c�a�ɍX�V
    Call dbcImportEdit.UpdateRecord
    'Call pErrorCheck
    Call pLockedControl(False)
    Call lblCIERROR_Change
    Call pButtonControl(False)
    Call sttausCIERROR  '//2014/05/19 �f�[�^���C�x���g���ɂ��낢�딭������̂ł����ɓ���
End Sub

Public Sub cmdEnd_Click()
    If True = pCheckEditData Then
        Dim stts As Integer
        stts = MsgBox("���e���ύX����Ă��܂��B" & vbCrLf & vbCrLf & "�X�V���܂����H", vbYesNoCancel + vbInformation, mCaption)
        Select Case stts
        Case vbYes
            Call cmdUpdate_Click
            If False = mUpdateOK Then
                Exit Sub
            End If
        Case vbNo
            Call cmdCancel_Click
        Case Else
            Exit Sub
        End Select
    End If
    Call dbcImportEdit.UpdateControls
    Call frmFurikaeReqImport.Show  '//�����I�ɔ�ь��̉�ʂ�\��
    Unload Me
End Sub

Private Sub cmdNext_Click()
    mIsActivated = False
    If True = pCheckEditData Then
        Dim stts As Integer
        stts = MsgBox("���e���ύX����Ă��܂��B" & vbCrLf & vbCrLf & "�X�V���܂����H", vbYesNoCancel + vbInformation, mCaption)
        Select Case stts
        Case vbYes
            Call cmdUpdate_Click
            If False = mUpdateOK Then
                Exit Sub
            End If
        Case vbNo
            Call cmdCancel_Click
        Case Else
            Exit Sub
        End Select
    End If
    Call gdDBS.MoveRecords(dbcImportEdit, 1)
    '//���C���� SpreadSheet �ɓ��e�𔽉f����FUpdate��ł� DataChanged() ���ω����Ă��܂��̂�...�B
    frmFurikaeReqImport.mEditRow = frmFurikaeReqImport.mEditRow + 1
    '//���ꂩ��ҏW����̂Ɋ��ɕҏW�ς݂ƂȂ��Ă���̂��������
    Call mForm.ResetDataControlEditFlag(Me)
    mErrMsgOn = False
    Call txtCIKYCD_KeyDown(vbKeyReturn, 0)
    mErrMsgOn = True
'    cmdUpdate.Enabled = False
'    cmdCancel.Enabled = False
    'Call dbcImportEdit.UpdateControls
    Call pButtonControl(False, True)
End Sub

Private Sub cmdPrev_Click()
    mIsActivated = False
    If True = pCheckEditData Then
        Dim stts As Integer
        stts = MsgBox("���e���ύX����Ă��܂��B" & vbCrLf & vbCrLf & "�X�V���܂����H", vbYesNoCancel + vbInformation, mCaption)
        Select Case stts
        Case vbYes
            Call cmdUpdate_Click
            If False = mUpdateOK Then
                Exit Sub
            End If
        Case vbNo
            Call cmdCancel_Click
        Case Else
            Exit Sub
        End Select
    End If
    Call gdDBS.MoveRecords(dbcImportEdit, -1)
    '//���C���� SpreadSheet �ɓ��e�𔽉f����FUpdate��ł� DataChanged() ���ω����Ă��܂��̂�...�B
    frmFurikaeReqImport.mEditRow = frmFurikaeReqImport.mEditRow - 1
    '//���ꂩ��ҏW����̂Ɋ��ɕҏW�ς݂ƂȂ��Ă���̂��������
    Call mForm.ResetDataControlEditFlag(Me)
    mErrMsgOn = False
    Call txtCIKYCD_KeyDown(vbKeyReturn, 0)
    mErrMsgOn = True
'    cmdUpdate.Enabled = False
'    cmdCancel.Enabled = False
    'Call dbcImportEdit.UpdateControls
    Call pButtonControl(False, True)
End Sub

Private Sub cmdKakutei_Click()
    If dblBankList.Text = "" Or dblShitenList.Text = "" Then
        Exit Sub
    End If
    txtCiBANK.Text = Left(dblBankList.Text, 4)
    txtCiSITN.Text = Left(dblShitenList.Text, 3)
    '//���͂��ꂽ���Z�@�֖����x�X����������������
    txtCiBKNM.Text = Mid(dblBankList.Text, 6)
    lblBankName.Caption = Mid(dblBankList.Text, 6)
    txtCiSINM.Text = Mid(dblShitenList.Text, 5)
    lblShitenName.Caption = Mid(dblShitenList.Text, 5)
    cmdKakutei.Enabled = False
'//2006/08/22 �m����M�\�ɁI
    Call pLockedControl(True)
End Sub

'///////////////////////////////////////////////////////
'//���R�[�h�ړ����ɂ��̃C�x���g���N����F�ҏW���J�n
Private Sub dbcImportEdit_Reposition()
    cmdNext.Enabled = Not dbcImportEdit.Recordset.IsLast
    cmdPrev.Enabled = Not dbcImportEdit.Recordset.IsFirst
    If dbcImportEdit.Recordset.BOF _
    Or dbcImportEdit.Recordset.EOF Then
        '//�擪�ȑO�A�Ō�ȍ~�̃��R�[�h�ʒu�͕ҏW�J�n�����Ȃ�
        Exit Sub
    End If
    'Debug.Print dbcImportEdit.Recordset.RowPosition
    '//�e���͍��ڂ̃G���[�\��
    Dim obj As Object
    For Each obj In Controls
        If TypeOf obj Is imText _
        Or TypeOf obj Is imNumber _
        Or TypeOf obj Is imDate Then
            If "" <> obj.DataField Then
                '//�S���� ORADC �Ƀo�C���h����Ă���͂��I
                obj.BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(obj.DataField & "E"))
            End If
        End If
    Next obj
    '//�ϑ��҃R�[�h�̃G���[�\��
    cboABKJNM.BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCiITKB.DataField & "E"))
    '//���Z�@�֋敪�̃G���[�\��
    optCiKKBN(0).BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCiKKBN.DataField & "E"), False)
    optCiKKBN(1).BackColor = optCiKKBN(0).BackColor
    '//�a����ʂ̃G���[�\��
    optCiKZSB(0).BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCiKZSB.DataField & "E"), False)
    optCiKZSB(1).BackColor = optCiKZSB(0).BackColor
    optCiKZSB(2).BackColor = optCiKZSB(0).BackColor
    cboCIOKFG.ListIndex = Val(lblCIOKFG.Caption) + 2    '// -2 �` 2
    chkCIMUPD.Value = Abs(Val(lblCIMUPD.Caption) <> 0)
    
    Call sttausCIERROR  '//2014/05/19 �f�[�^���C�x���g���ɂ��낢�딭������̂ł����ɓ���
    
    Call dbcImportEdit.Recordset.Edit         '//�ҏW�J�n
End Sub

Private Sub dblBankList_Click()
    cboShitenYomi.ListIndex = -1
    Call cboShitenYomi_Click
End Sub

Private Sub dblShitenList_Click()
    cmdKakutei.Enabled = dblBankList.Text <> ""
End Sub

Private Sub Form_Activate()
    mCheckUpdate = True     '//�X�V�{�^���̐���F�f�[�^�\�����ɃC�x���g���������Ă��\�Ȃ悤�ɁI
    If False = mIsActivated Then
        Call pButtonControl(False, True)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
    mErrMsgOn = True
    '//�L�[�����������ɍX�V�\�����f
'    cmdUpdate.Enabled = pCheckEditData
'    cmdCancel.Enabled = cmdUpdate.Enabled
End Sub

Private Sub Form_Load()
    mCheckUpdate = False    '//�X�V�{�^���̐���F�f�[�^�\�����ɃC�x���g���������Ă��\�Ȃ悤�ɁI
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    Call mForm.MoveSysDate
    '//��s�ƗX�֋ǂ� Frame �𐮗񂷂�
    fraBank(0).ZOrder
    fraBank(1).Top = fraBank(0).Top
    fraBank(1).Left = fraBank(0).Left
    fraBank(1).Height = fraBank(0).Height
    fraBank(1).Width = fraBank(0).Width
    fraBank(0).BorderStyle = vbBSNone
    fraBank(1).BorderStyle = vbBSNone
    fraBankList.BorderStyle = vbBSNone
    cmdKakutei.Enabled = False
    imgCIWMSG.Visible = False
    lblCIWMSG.Caption = ""
    lblCIWMSG.Visible = False
    'lblCIWMSG.AutoSize = True
    Call mRimp.UpdateComboBox(cboCIOKFG)
 
    dbcBank.RecordSource = ""
    dbcShiten.RecordSource = ""
    '//�Ăяo�����Őݒ肷��̂ŕs�v
    'dbcImportEdit.RecordSource = frmFurikaeReqImport.dbcImport.RecordSource
    dbcItakushaMaster.RecordSource = "SELECT * FROM taItakushaMaster ORDER BY ABITCD"
    dbcItakushaMaster.ReadOnly = True
    Call pLockedControl(False)
    Call mForm.pInitControl
    lblBAKJNM.Caption = ""
    lblBankName.Caption = ""
    lblShitenName.Caption = ""
    Call gdDBS.SetItakushaComboBox(cboABKJNM)
    'Call cmdEnd.SetFocus
    Call pButtonControl(False)
End Sub

Private Sub Form_Resize()
    '//����ȏ㏬��������ƃR���g���[�����B���̂Ő��䂷��
    If Me.Height < 8000 Then
        Me.Height = 8000
    End If
    If Me.Width < 10200 Then
        Me.Width = 10200
    End If
    Call mForm.Resize
    Call mForm.MoveSysDate
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFurikaeReqImportEdit = Nothing
    '//�q�t�H�[���Ƃ��đ��݂���̂�j��
    Set gdFormSub = Nothing
    Set mForm = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

'//2014/05/19 �f�[�^���C�x���g���ɂ��낢�딭������̂ł����ɓ���
Private Sub sttausCIERROR()
    Dim err As Integer
    err = Val(lblCIERROR.Caption)
    If err = mRimp.errInvalid And 0 <> Val(lblCIMUPD.Caption) Then
        err = mRimp.errWarning
    End If
    Select Case err
    Case mRimp.errImport:     lblERRMSG.Caption = "�捞": lblERRMSG.BackColor = mRimp.ErrorStatus(err)
    Case mRimp.errEditData:   lblERRMSG.Caption = "�C��": lblERRMSG.BackColor = mRimp.ErrorStatus(err)
    Case mRimp.errInvalid:    lblERRMSG.Caption = "�ُ�": lblERRMSG.BackColor = mRimp.ErrorStatus(err)
    Case mRimp.errNormal:     lblERRMSG.Caption = "����": lblERRMSG.BackColor = vbCyan
    Case mRimp.errWarning:    lblERRMSG.Caption = "�x��": lblERRMSG.BackColor = mRimp.ErrorStatus(err)
    Case Else:                lblERRMSG.Caption = "��O": lblERRMSG.BackColor = vbRed
    End Select
    'lblERRMSG.BackColor = mRimp.ErrorStatus(lblCIERROR.Caption)
    '//2014/05/19 �X�V���[�h�̒ǉ�
    If err = mRimp.errInvalid Then
        '//�ُ�f�[�^���ɂ͎g�p�ł��Ȃ��悤�ɐ��䂷��
        fraCiINSD.Enabled = False
    Else
        '//�ی�҃}�X�^�Ƀf�[�^���������ɂ͎g�p�ł��Ȃ��悤�ɐ��䂷��
        fraCiINSD.Enabled = checkExists()
    End If
    lblUpdMode.ForeColor = IIf(fraCiINSD.Enabled, vbBlue, vbButtonShadow)
    optCiINSD(0).ForeColor = IIf(fraCiINSD.Enabled, vbButtonText, vbButtonShadow)
    optCiINSD(1).ForeColor = IIf(fraCiINSD.Enabled, vbButtonText, vbButtonShadow)
End Sub

Private Sub lblCIERROR_Change()
    'Call sttausCIERROR
End Sub

Private Sub lblCIKKBN_Change()
'    On Error Resume Next
    '//�u�����N�̓G���[�Ƃ���
    If Not IsNull(lblCiKKBN.Caption) And "" <> lblCiKKBN.Caption Then
        optCiKKBN(lblCiKKBN.Caption).Value = True
    End If
End Sub

Private Sub lblCIKZSB_Change()
    If Not IsNull(lblCiKZSB.Caption) And "" <> lblCiKZSB.Caption Then
        optCiKZSB(Val(lblCiKZSB.Caption)).Value = True
    Else
'//�ݒ肷��ƍX�V�t���O�������Ă��܂��̂Ŏ~�߂�
'//        optCIKZSB(0).Value = True
    End If
End Sub

Private Sub lblCIWMSG_Change()
    txtCIWMSG.Text = lblCIWMSG.Caption
    imgCIWMSG.Visible = lblCIWMSG.Caption <> ""
End Sub

Private Sub optCIKKBN_Click(Index As Integer)
    fraKinnyuuKikan.Tag = Index
    Call fraBank(Index).ZOrder(0)
    fraBankList.Visible = Index = 0
    lblCiKKBN.Caption = Index
    '//�t�H�[�J�X��������̂Őݒ肷��.
    txtCiBANK.TabStop = Index = eBankKubun.KinnyuuKikan
    txtCiSITN.TabStop = Index = eBankKubun.KinnyuuKikan
    txtCiKZNO.TabStop = Index = eBankKubun.KinnyuuKikan
    txtCiYBTK.TabStop = Index = eBankKubun.YuubinKyoku
    txtCiYBTN.TabStop = Index = eBankKubun.YuubinKyoku
'    cmdUpdate.Enabled = True
'    cmdCancel.Enabled = True
    Call pButtonControl(True)
End Sub

Private Sub optCIKZSB_Click(Index As Integer)
    lblCiKZSB.Caption = Index
'    cmdUpdate.Enabled = True
'    cmdCancel.Enabled = True
    Call pButtonControl(True)
End Sub

Private Sub optCiINSD_Click(Index As Integer)
    fraCiINSD.Tag = Index
    lblCiINSD.Caption = Index
    Call pButtonControl(True)
End Sub

Private Sub lblCiINSD_Change()
'    On Error Resume Next
    '//�u�����N�̓G���[�Ƃ���
    If Not IsNull(lblCiINSD.Caption) And "" <> lblCiINSD.Caption Then
        optCiINSD(Abs(lblCiINSD.Caption)).Value = True
    End If
End Sub

Private Sub txtCIBANK_Change()
    Call pButtonControl(True)
End Sub
Private Sub txtCiBANK_LostFocus()
    If 0 <= Len(Trim(txtCiBANK.Text)) And Len(Trim(txtCiBANK.Text)) < 4 Then
        lblBankName.Caption = ""
        Exit Sub
    End If
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Set dyn = gdDBS.SelectBankMaster("DISTINCT DAKJNM", eBankRecordKubun.Bank, Trim(txtCiBANK.Text), vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblBankName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))
End Sub

Private Sub txtCIBKNM_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCiFKxx_Change(Index As Integer)
    Call pButtonControl(True)
End Sub

Private Sub txtCIHGCD_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCiKJNM_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCiKNNM_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCIKZNM_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCIKZNO_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCISINM_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCISITN_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCIKYCD_Change()
    Call pButtonControl(True)
End Sub

Public Sub txtCIKYCD_KeyDown(KeyCode As Integer, Shift As Integer)
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
        
'//2013/06/18 �O�[�����ߍ���
    txtCiKYCD.Text = Format(Val(txtCiKYCD.Text), String(7, "0"))
''    If "" = Trim(txtCiKYCD.Text) Then
''        Exit Sub
''    End If
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//    sql = "SELECT DISTINCT BAITKB,BAKYCD,BAKSCD,BAKJNM FROM tbKeiyakushaMaster"
'//2015/02/12 �ŐV�f�[�^�P���݂̂Ŕ���
'//    sql = "SELECT DISTINCT BAITKB,BAKYCD,BAKJNM,BAKYED FROM tbKeiyakushaMaster"
    sql = "SELECT BAITKB,BAKYCD,BAKJNM,BAKYED FROM tbKeiyakushaMaster"
    sql = sql & " WHERE BAITKB = '" & lblCiITKB.Caption & "'"
    sql = sql & "   AND BAKYCD = '" & txtCiKYCD.Text & "'"
'//2006/03/31 ����Ԃ�\������悤�ɕύX
'    sql = sql & "   AND TO_CHAR(SYSDATE,'YYYYMMDD') BETWEEN BAKYST AND BAKYED" '//�L���f�[�^�i����
'//2015/02/12 �ŐV�f�[�^�P���݂̂Ŕ���
    sql = sql & " ORDER BY BASQNO DESC"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If 0 = dyn.RecordCount Then
        Call dyn.Close
        KeyCode = 0
        lblBAKJNM.Caption = ""
        If mErrMsgOn = True Then
            '//                                        �u�_��Ҕԍ��v
            Call MsgBox("�Y���f�[�^�͑��݂��܂���.(" & lblKeiyakushaCode.Caption & ")", vbInformation, mCaption)
            Call txtCiKYCD.SetFocus
        End If
        'Exit Sub
    Else
#If 1 Then
'//2015/02/12 ����Ԃ�\������悤�ɕύX
        lblBAKJNM.ForeColor = vbBlack
        Dim kk As String
        If dyn.Fields("BAKYED") < gdDBS.sysDate("yyyymmdd") Then
            kk = "(���)"
            lblBAKJNM.ForeColor = vbRed
        End If
        lblBAKJNM.Caption = kk & dyn.Fields("BAKJNM")
#Else
        lblBAKJNM.Caption = IIf(dyn.Fields("BAKYED") < gdDBS.sysDate("yyyymmdd"), "(���)", "") & _
                            dyn.Fields("BAKJNM")
#End If
    End If

#If 0 Then
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
    Call cboCIKSCDz.Clear
    Do Until dyn.EOF
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//        Call cboCIKSCDz.AddItem(dyn.Fields("BAKSCD"))
        Call dyn.MoveNext
    Loop
    cboCIKSCDz.ListIndex = 0
#End If
    Call dyn.Close
    '//2007/06/06   ��s���E�x�X���̓ǂݍ��݂������ł���悤�ɕύX
    '//             �Ǎ��ݎ��� Change()=���̕\�� �C�x���g���Ԃ� �x�X�R�[�h�E��s�R�[�h�̏��ɂȂ�x�X�����\������Ȃ����Ƃ�����
    Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Bank, txtCiBANK.Text, vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblBankName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))
    Set dyn = Nothing
    Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Shiten, txtCiBANK.Text, txtCiSITN.Text, vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"�x�X��_����" �œǂ߂Ȃ�
    Set dyn = Nothing
    'txtCIKJNM.SetFocus
End Sub

Private Function pUpdateErrorCheck() As Boolean
'//2012/07/11 �}�X�^���f���Ȃ��ꍇ�`�F�b�N���Ȃ�
    If chkCIMUPD.Value <> 0 Then
        pUpdateErrorCheck = True
        Exit Function
    End If
'//2006/06/26 �X�V���̃`�F�b�N���Ȃ������̂Œǉ��F�ی�҃����e���R�s�[
    '///////////////////////////////
    '//�K�{���͍��ڂƐ������`�F�b�N
    
    Dim str As New StringClass
    Dim obj As Object, msg As String
    '//�ی�ҁE�������͕̂K�{
    If txtCiKJNM.Text = "" Then
        Set obj = txtCiKJNM
        msg = "�ی�Җ�(����)�͕K�{���͂ł�."
    ElseIf False = str.CheckLength(txtCiKJNM.Text) Then
        Set obj = txtCiKJNM
        msg = "�ی�Җ�(����)�ɔ��p���܂܂�Ă��܂�."
    End If
    '//�ی�ҁE�J�i���͕̂K�{
    '//2007/06/07 �K�{ �����F�������`�l�Ɠ����l�Ƃ����
    If txtCiKNNM.Text = "" Then
        Set obj = txtCiKNNM
        msg = "�ی�Җ�(�J�i)�͕K�{���͂ł�."
    ElseIf False = str.CheckLength(txtCiKNNM.Text, vbNarrow) Then
        Set obj = txtCiKNNM
        msg = "�ی�Җ�(�J�i)�ɑS�p���܂܂�Ă��܂�."
    ElseIf 0 < InStr(txtCiKNNM.Text, "�") Then
        Set obj = txtCiKNNM
        msg = "�ی�Җ�(�J�i)�ɒ������܂܂�Ă��܂�."
    End If
#If 0 Then  '//���ڂȂ�
    If IsNull(txtCIKYxx(1).Number) Then
        Set obj = txtCIKYxx(1)
        msg = "�_����Ԃ̏I�����͕K�{���͂ł�."
    ElseIf txtCIKYxx(0).Text > txtCIKYxx(1).Text Then
        Set obj = txtCIKYxx(0)
        msg = "�_����Ԃ��s���ł�."
    ElseIf IsNull(txtCiFKxx(1).Number) Then
        Set obj = txtCiFKxx(1)
        msg = "�U�֊��Ԃ̏I�����͕K�{���͂ł�."
    ElseIf txtCiFKxx(0).Text > txtCiFKxx(1).Text Then
        Set obj = txtCiFKxx(0)
        msg = "�U�֊��Ԃ��s���ł�."
    End If
#End If
    If lblCiKKBN.Caption = "" Then
        If txtCiBANK.Visible = True And txtCiBANK.Enabled = True Then
            Set obj = txtCiBANK
        ElseIf txtCiYBTK.Visible = True And txtCiYBTK.Enabled = True Then
            Set obj = txtCiYBTK
        Else
            Set obj = txtCiKYCD     '// ==> fraKinnyuuKikan �ɂ̓t�H�[�J�X�𓖂Ă��Ȃ��̂ł���������
        End If
        msg = "���Z�@�֋敪�͕K�{���͂ł�."
    ElseIf lblCiKKBN.Caption = eBankKubun.KinnyuuKikan Then
        If txtCiBANK.Text = "" Or lblBankName.Caption = "" Then
            Set obj = txtCiBANK
            msg = "���Z�@�ւ͕K�{���͂ł�."
        ElseIf txtCiSITN.Text = "" Or lblShitenName.Caption = "" Then
            Set obj = txtCiSITN
            msg = "�x�X�͕K�{���͂ł�."
        ElseIf Not (lblCiKZSB.Caption = eBankYokinShubetsu.Futsuu _
                 Or lblCiKZSB.Caption = eBankYokinShubetsu.Touza) Then
            Set obj = optCiKZSB(eBankYokinShubetsu.Futsuu)
            msg = "�a����ʂ͕K�{���͂ł�."
        ElseIf txtCiKZNO.Text = "" Then
            Set obj = txtCiKZNO
            msg = "�����ԍ��͕K�{���͂ł�."
        End If
    ElseIf lblCiKKBN.Caption = eBankKubun.YuubinKyoku Then
        If txtCiYBTK.Text = "" Then
            Set obj = txtCiYBTK
            msg = "�ʒ��L���͕K�{���͂ł�."
        ElseIf txtCiYBTN.Text = "" Then
            Set obj = txtCiYBTN
            msg = "�ʒ��ԍ��͕K�{���͂ł�."
        ElseIf "1" <> Right(txtCiYBTN.Text, 1) Then
'//2006/04/26 �����ԍ��`�F�b�N
            Set obj = txtCiYBTN
            msg = "�ʒ��ԍ��̖������u�P�v�ȊO�ł�."
        End If
    End If
    If txtCiKZNM.Text = "" Then
        Set obj = txtCiKZNM
        msg = "�������`�l(�J�i)�͕K�{���͂ł�."
    End If
    '//Object ���ݒ肳��Ă��邩�H
    If TypeName(obj) <> "Nothing" Then
        Call MsgBox(msg, vbCritical, mCaption)
        Call obj.SetFocus
        Exit Function
    End If
    pUpdateErrorCheck = True
    Exit Function
pUpdateErrorCheckError:
    Call gdDBS.ErrorCheck       '//�G���[�g���b�v
    pUpdateErrorCheck = False   '//���S�̂��߁FFalse �ŏI������͂�
End Function

Private Sub mnuEnd_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

Private Sub txtCISITN_LostFocus()
    If 0 <= Len(Trim(txtCiSITN.Text)) And Len(Trim(txtCiSITN.Text)) < 3 Then
        lblShitenName.Caption = ""
        Exit Sub
    End If
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Shiten, Trim(txtCiBANK.Text), Trim(txtCiSITN.Text), vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"�x�X��_����" �œǂ߂Ȃ�
End Sub

Private Sub txtCIYBTK_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCIYBTN_Change()
    Call pButtonControl(True)
End Sub

'/////////////////////////
'//�ăG���[�`�F�b�N�čl�I�I�I���R�[�h�̍X�V���o���Ȃ��Ȃ�
Private Function pErrorCheck()
    '//�e���͍��ڂ̃G���[�\��
    Dim obj As Object
    
    Call frmFurikaeReqImport.gDataCheck(Format(lblCIINDT.Caption, "yyyy/MM/dd hh:nn:ss"), lblCISEQN.Caption)
    For Each obj In Controls
        If TypeOf obj Is imText _
        Or TypeOf obj Is imNumber _
        Or TypeOf obj Is imDate Then
            If "" <> obj.DataField Then
                '//�S���� ORADC �Ƀo�C���h����Ă���͂��I
                obj.BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(obj.DataField & "E"))
            End If
        End If
    Next obj
    '//�ϑ��҃R�[�h�̃G���[�\��
    cboABKJNM.BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCiITKB.DataField & "E"))
    '//���Z�@�֋敪�̃G���[�\��
    optCiKKBN(0).BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCiKKBN.DataField & "E"), False)
    optCiKKBN(1).BackColor = optCiKKBN(0).BackColor
    '//�a����ʂ̃G���[�\��
    optCiKZSB(0).BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCiKZSB.DataField & "E"), False)
    optCiKZSB(1).BackColor = optCiKZSB(0).BackColor
    optCiKZSB(2).BackColor = optCiKZSB(0).BackColor
End Function

'//�ی�҃}�X�^�̃��R�[�h�����ɑ��݂��邩
Private Function checkExists()
    checkExists = InStr(lblCIWMSG.Caption, MainModule.cEXISTS_DATA) <> 0 _
               Or InStr(lblCIWMSG.Caption, MainModule.cKAIYAKU_DATA) <> 0
End Function

'//�ی�҃}�X�^�̃��R�[�h�����݂��A����Ԃł��邩
Private Function checkKaiyaku()
    checkKaiyaku = InStr(lblCIWMSG.Caption, MainModule.cKAIYAKU_DATA) <> 0
End Function

