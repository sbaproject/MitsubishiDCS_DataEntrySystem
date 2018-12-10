VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHogoshaMaster 
   Caption         =   "保護者マスタメンテナンス"
   ClientHeight    =   7335
   ClientLeft      =   1710
   ClientTop       =   4725
   ClientWidth     =   10125
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
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
      Caption         =   "教室番号(&Z)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      ToolTipText     =   "前後の履歴に移動"
      Top             =   1860
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.ComboBox cboABKJNM 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "保護者マスタメンテナンス.frx":0000
      Left            =   1800
      List            =   "保護者マスタメンテナンス.frx":000D
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "InputKey"
      Top             =   900
      Width           =   1755
   End
   Begin VB.Frame fraKinnyuuKikan 
      Caption         =   "振替口座"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "郵便局"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
            Caption         =   "保護者マスタメンテナンス.frx":002B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "保護者マスタメンテナンス.frx":0097
            Key             =   "保護者マスタメンテナンス.frx":00B5
            MouseIcon       =   "保護者マスタメンテナンス.frx":00F9
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
            Caption         =   "保護者マスタメンテナンス.frx":0115
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "保護者マスタメンテナンス.frx":0181
            Key             =   "保護者マスタメンテナンス.frx":019F
            MouseIcon       =   "保護者マスタメンテナンス.frx":01E3
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
            Alignment       =   1  '右揃え
            Caption         =   "通帳番号"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
            Alignment       =   1  '右揃え
            Caption         =   "通帳記号"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "民間金融機関"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
            Caption         =   "保護者マスタメンテナンス.frx":01FF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "保護者マスタメンテナンス.frx":026B
            Key             =   "保護者マスタメンテナンス.frx":0289
            MouseIcon       =   "保護者マスタメンテナンス.frx":02CD
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
            Caption         =   "保護者マスタメンテナンス.frx":02E9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "保護者マスタメンテナンス.frx":0355
            Key             =   "保護者マスタメンテナンス.frx":0373
            MouseIcon       =   "保護者マスタメンテナンス.frx":03B7
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
            Caption         =   "保護者マスタメンテナンス.frx":03D3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "保護者マスタメンテナンス.frx":043F
            Key             =   "保護者マスタメンテナンス.frx":045D
            MouseIcon       =   "保護者マスタメンテナンス.frx":04A1
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
            BorderStyle     =   0  'なし
            Caption         =   "口座種別"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
               Caption         =   "当座"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
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
               Caption         =   "普通"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
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
                  Name            =   "ＭＳ Ｐゴシック"
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
               Caption         =   "口座種別"
               DataField       =   "CAKZSB"
               DataSource      =   "dbcHogoshaMaster"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
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
            Caption         =   "東京三菱５６７x"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
            Caption         =   "取引銀行"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
            Caption         =   "取引支店"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
            Caption         =   "口座種別"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
            Caption         =   "口座番号"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
            Caption         =   "大阪３４５６７x"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "郵便局"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "民間金融機関"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "保護者マスタメンテナンス.frx":04BD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "保護者マスタメンテナンス.frx":0529
         Key             =   "保護者マスタメンテナンス.frx":0547
         MouseIcon       =   "保護者マスタメンテナンス.frx":058B
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
         Text            =   "ｺｳｻﾞﾒｲｷﾞﾆﾝﾒｲ...........................*"
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
         Alignment       =   1  '右揃え
         Caption         =   "口座名義人(カナ)"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "金融機関種別"
         DataField       =   "CAKKBN"
         DataSource      =   "dbcHogoshaMaster"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "解約"
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
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "保護者マスタメンテナンス.frx":05A7
      Left            =   2880
      List            =   "保護者マスタメンテナンス.frx":05B4
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
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
      Calculator      =   "保護者マスタメンテナンス.frx":05D2
      Caption         =   "保護者マスタメンテナンス.frx":05F2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタメンテナンス.frx":065E
      Keys            =   "保護者マスタメンテナンス.frx":067C
      MouseIcon       =   "保護者マスタメンテナンス.frx":06C6
      Spin            =   "保護者マスタメンテナンス.frx":06E2
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
      Caption         =   "金融機関リスト"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "保護者マスタメンテナンス.frx":070A
         Left            =   1500
         List            =   "保護者マスタメンテナンス.frx":072F
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   180
         Width           =   855
      End
      Begin VB.ComboBox cboShitenYomi 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "保護者マスタメンテナンス.frx":0771
         Left            =   3900
         List            =   "保護者マスタメンテナンス.frx":0796
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdKakutei 
         Caption         =   "確定(&K)"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         RecordSource    =   ""
      End
      Begin MSDBCtls.DBList dblBankList 
         Bindings        =   "保護者マスタメンテナンス.frx":07D8
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
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDBCtls.DBList dblShitenList 
         Bindings        =   "保護者マスタメンテナンス.frx":07EE
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
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label24 
         Caption         =   "金融機関 読み⇒"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "支店　　　　読み⇒"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "中止(&C)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "処理区分"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "参照"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "修正"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "削除"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "新規"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Caption         =   "処理区分"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "更新(&U)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "終了(&X)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Calendar        =   "保護者マスタメンテナンス.frx":0806
      Caption         =   "保護者マスタメンテナンス.frx":0986
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタメンテナンス.frx":09F2
      Keys            =   "保護者マスタメンテナンス.frx":0A10
      MouseIcon       =   "保護者マスタメンテナンス.frx":0A6E
      Spin            =   "保護者マスタメンテナンス.frx":0A8A
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
      Calendar        =   "保護者マスタメンテナンス.frx":0AB2
      Caption         =   "保護者マスタメンテナンス.frx":0C32
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタメンテナンス.frx":0C9E
      Keys            =   "保護者マスタメンテナンス.frx":0CBC
      MouseIcon       =   "保護者マスタメンテナンス.frx":0D1A
      Spin            =   "保護者マスタメンテナンス.frx":0D36
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
      Calendar        =   "保護者マスタメンテナンス.frx":0D5E
      Caption         =   "保護者マスタメンテナンス.frx":0EDE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタメンテナンス.frx":0F4A
      Keys            =   "保護者マスタメンテナンス.frx":0F68
      MouseIcon       =   "保護者マスタメンテナンス.frx":0FC6
      Spin            =   "保護者マスタメンテナンス.frx":0FE2
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
      Caption         =   "保護者マスタメンテナンス.frx":100A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタメンテナンス.frx":1076
      Key             =   "保護者マスタメンテナンス.frx":1094
      MouseIcon       =   "保護者マスタメンテナンス.frx":10D8
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
      Text            =   "漢字氏名．．．．．．．．．．＊"
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
      Caption         =   "保護者マスタメンテナンス.frx":10F4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタメンテナンス.frx":1160
      Key             =   "保護者マスタメンテナンス.frx":117E
      MouseIcon       =   "保護者マスタメンテナンス.frx":11C2
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
      Caption         =   "保護者マスタメンテナンス.frx":11DE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタメンテナンス.frx":124A
      Key             =   "保護者マスタメンテナンス.frx":1268
      MouseIcon       =   "保護者マスタメンテナンス.frx":12AC
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
      Calculator      =   "保護者マスタメンテナンス.frx":12C8
      Caption         =   "保護者マスタメンテナンス.frx":12E8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタメンテナンス.frx":1354
      Keys            =   "保護者マスタメンテナンス.frx":1372
      MouseIcon       =   "保護者マスタメンテナンス.frx":13BC
      Spin            =   "保護者マスタメンテナンス.frx":13D8
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
      Caption         =   "保護者マスタメンテナンス.frx":1400
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタメンテナンス.frx":146C
      Key             =   "保護者マスタメンテナンス.frx":148A
      MouseIcon       =   "保護者マスタメンテナンス.frx":14CE
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
      Text            =   "生徒氏名．．．．．．．．．．＊"
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
      Caption         =   "保護者マスタメンテナンス.frx":14EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタメンテナンス.frx":1556
      Key             =   "保護者マスタメンテナンス.frx":1574
      MouseIcon       =   "保護者マスタメンテナンス.frx":15B8
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
      Text            =   "ｶﾅｼﾒｲ..................................*"
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
      Calendar        =   "保護者マスタメンテナンス.frx":15D4
      Caption         =   "保護者マスタメンテナンス.frx":1754
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタメンテナンス.frx":17C0
      Keys            =   "保護者マスタメンテナンス.frx":17DE
      MouseIcon       =   "保護者マスタメンテナンス.frx":183C
      Spin            =   "保護者マスタメンテナンス.frx":1858
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
      Caption         =   "保護者マスタメンテナンス.frx":1880
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "保護者マスタメンテナンス.frx":18EC
      Key             =   "保護者マスタメンテナンス.frx":190A
      MouseIcon       =   "保護者マスタメンテナンス.frx":194E
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
      Alignment       =   1  '右揃え
      Caption         =   "保護者名(カナ)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "生徒氏名"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "変更後金額"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "作成日"
      DataField       =   "CAADDT"
      DataSource      =   "dbcHogoshaMaster"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "解約フラグ"
      DataField       =   "CAKYFG"
      DataSource      =   "dbcHogoshaMaster"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "振替終了日"
      DataField       =   "CAFKED"
      DataSource      =   "dbcHogoshaMaster"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "振替開始日"
      DataField       =   "CAFKST"
      DataSource      =   "dbcHogoshaMaster"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "契約終了日"
      DataField       =   "CAKYED"
      DataSource      =   "dbcHogoshaMaster"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "契約開始日"
      DataField       =   "CAKYST"
      DataSource      =   "dbcHogoshaMaster"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "更新者"
      DataField       =   "CAUSID"
      DataSource      =   "dbcHogoshaMaster"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "更新日"
      DataField       =   "CAUPDT"
      DataSource      =   "dbcHogoshaMaster"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "教室番号"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "委託者区分"
      DataField       =   "CAITKB"
      DataSource      =   "dbcHogoshaMaster"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "教室番号"
      DataField       =   "CAKSCD"
      DataSource      =   "dbcHogoshaMaster"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "委託者区分"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "保護者ＳＥＱ"
      DataField       =   "CASQNO"
      DataSource      =   "dbcHogoshaMaster"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "契約者番号"
      DataField       =   "CAKYCD"
      DataSource      =   "dbcHogoshaMaster"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "保護者番号"
      DataField       =   "CAHGCD"
      DataSource      =   "dbcHogoshaMaster"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "請求予定額"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "田中　俊彦"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "契約者番号"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "保護者番号"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "保護者名(漢字)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "口座振替期間"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "契約期間"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
Attribute VB_Name = "frmHogoshaMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm As New FormClass
Private mCaption As String
Private mIsActivated As Boolean
'//2013/02/26 口座変更等の更新時の追加更新の際に２度 pUpdateRecord() が実行されるのを制御する
Private mRirekiAddNewUpdate As Boolean

'//2007/06/07 更新・中止ボタンを完全単独にコントロール
Private Sub pButtonControl(ByVal vMode As Boolean, Optional vExec As Boolean = False)
    If True = mIsActivated Or True = vExec Then
        cmdUpdate.Visible = vMode
        cmdCancel.Visible = vMode
        cmdUpdate.Enabled = vMode
        cmdCancel.Enabled = vMode
        cmdEnd.Enabled = Not vMode
        mIsActivated = True
    End If
    '//修正時以外は教室番号ボタンの押下は不可能にする
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
    '//dblBankList.Refresh() を実行すると下は不要
'    cboShitenYomi.ListIndex = -1
'    dblShitenList.ListField = ""
'    dblShitenList.Refresh
    Call dblShitenList.ReFill
    'cmdEnd.Enabled = blMode
    spnRireki.Visible = False
    '//2007/06/07 口座名義人は常に入力しない：保護者名(カナ)をコピーする様に仕様変更
    txtCAKZNM.Enabled = False
    lblKouzaName.Enabled = False
    cmdKakutei.Enabled = Not blMode
End Sub

#If 0 Then
Private Sub cboCAKSCDz_GotFocus()
    '//TAB キー入力時 txtCAKYCD_KeyDown() イベントが発生しないので
    Call cboCAKYCDz_KeyDown(vbKeyReturn, 0)
End Sub
#End If

Private Sub chkCAKYFG_Click()
    lblCAKYFG.Caption = chkCAKYFG.Value
    Call pButtonControl(True)
End Sub

Private Sub chkCAKYFG_KeyDown(KeyCode As Integer, Shift As Integer)
    '//解約フラグを設定したので終了日の入力を促す.
    '//KeyCode & Shift をクリアしないとバッファに残る？
    KeyCode = 0
    Shift = 0
    chkCAKYFG.Value = Choose(chkCAKYFG.Value + 1, 1, 0, 0)  '// Index=1,2,3
    Call MsgBox("解約の変更を検知しました。" & vbCrLf & vbCrLf & "契約期間及び振替期間 終了日の再設定をして下さい.", vbInformation + vbOKOnly, mCaption)
    Call txtCAKYxx(1).SetFocus
End Sub

Private Sub chkCAKYFG_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '//解約フラグを設定したので終了日の入力を促す.
    If Button = vbLeftButton Then
        Call chkCAKYFG_KeyDown(vbKeySpace, 0)
    End If
End Sub

Private Sub cmdClassNoChange_Click()
    Load frmClassNoChange
    With frmClassNoChange
        '//表示フォームをこのフォームの中央にする.
        .Top = Me.Top + (Me.Height - .Height) / 2
        .Left = Me.Left + (Me.Width - .Width) / 2
        .lblCAITKB.Caption = lblCAITKB.Caption
        .lblCAKYCD.Caption = lblCAKYCD.Caption
        .lblCAKSCD.Caption = lblCAKSCD.Caption
        .lblCAHGCD.Caption = lblCAHGCD.Caption
        .txtCAKSCD.Text = ""
        Call .Show(vbModal)
        '//変更されたので .mNewCode に値が入る
        If "" <> Trim(.mNewCode) Then
            txtCAKSCD.Text = .mNewCode
            'lblCAKSCD.Caption = .mNewCode
            '//排他がかかるので frmClassNoChange() で更新前に保護者マスタ・ロック解除している：更新前なら不要
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
'//2006/04/24 ここから：教室番号のユニーク性をチェック：教室番号はなぜユニークから外したか？
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
            Call MsgBox("既にデータが存在します.(" & lblHogoshaCode.Caption & ")", vbCritical, mCaption)
            Exit Function
        End If
    End If
'//2006/04/24 ここまで：教室番号のユニーク性をチェック：教室番号はなぜユニークから外したか？
'///////////////////////////////////////////////////////////////////////////////////////////
#End If

'''//2002/10/18 そのままの日付とする
'''    lblCAKYxx(0).Caption = gdDBS.FirstDay(txtCAKYxx(0).Number)
'''    lblCAKYxx(1).Caption = gdDBS.LastDay(txtCAKYxx(1).Number)
'''    lblCAFKxx(0).Caption = gdDBS.FirstDay(txtCAFKxx(0).Number)
'''    lblCAFKxx(1).Caption = gdDBS.LastDay(txtCAFKxx(1).Number)
    lblCAKYxx(0).Caption = Val(gdDBS.Nz(txtCAKYxx(0).Number))
    lblCAKYxx(1).Caption = Val(gdDBS.Nz(txtCAKYxx(1).Number))
    lblCAFKxx(0).Caption = Val(gdDBS.Nz(txtCAFKxx(0).Number))
    lblCAFKxx(1).Caption = Val(gdDBS.Nz(txtCAFKxx(1).Number))
'//2003/01/31 解約フラグが NULL になるので変更
    lblCAKYFG.Caption = Val(chkCAKYFG.Value)
    lblCAUSID.Caption = gdDBS.LoginUserName
    If "" = lblCAADDT.Caption Then
        lblCAADDT.Caption = gdDBS.sysDate
    End If
    lblCAUPDT.Caption = gdDBS.sysDate
    Call dbcHogoshaMaster.UpdateRecord
'//2004/07/09 口座振替データは旧のままにしておく：変更前・後の差異をとるため
#If 0 Then
    '//2003/01/31 口座振替予定データへの更新
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
'//2004/07/09 解約者の更新追加
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
            Call MsgBox("口座振替データで使用されているため" & vbCrLf & vbCrLf & "削除する事は出来ません.", vbCritical, mCaption)
            Exit Sub
        End If
        If vbOK <> MsgBox("削除しますか？" & vbCrLf & vbCrLf & "元に戻すことは出来ません.", vbInformation + vbOKCancel + vbDefaultButton2, mCaption) Then
            Exit Sub
        Else
'//2002/11/26 OIP-00000 ORA-04108 でエラーになるので Execute() で実行するように変更.
'// Oracle Data Control 8i(3.6) 9i(4.2) の違いかな？
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
'//2013/02/26 口座変更等の更新時の追加更新の際に２度 pUpdateRecord() が実行されるのを制御する
        mRirekiAddNewUpdate = False
        '//入力内容チェックで取りやめしたので終了
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
    cmdClassNoChange.Visible = False    '//教室番号修正不可！
End Sub

Private Sub cmdCancel_Click()
    Call dbcHogoshaMaster.UpdateControls
    Call pLockedControl(True)
    Call txtCAKYCD.SetFocus ' cboABKJNM.SetFocus
    Call pButtonControl(False)
    cmdClassNoChange.Visible = False    '//教室番号修正不可！
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
    '//銀行と郵便局の Frame を整列する
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
    '//初期値をセット
    optShoriKubun(0).Value = True
 
    dbcBank.RecordSource = ""
    dbcShiten.RecordSource = ""
    dbcHogoshaMaster.RecordSource = ""
    dbcItakushaMaster.RecordSource = "SELECT * FROM taItakushaMaster ORDER BY ABITCD"
    dbcItakushaMaster.ReadOnly = True
    Call pLockedControl(True)
    Call mForm.pInitControl
    '//契約者・保護者コード入力時はこの定義を外す
    'txtCAKYCD.KeyNext = ""
    'txtCAHGCD.KeyNext = ""
    '//初期値をセット：修正モード
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
    '//フォーカスが消えるので設定する.
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
    On Error Resume Next    'Form_Load()時にフォーカスを当てられない時エラーとなるので回避のエラー処理
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
    lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"支店名_漢字" で読めない
End Sub

Private Sub spnRireki_DownClick()
    '//後のレコードに移動
    If True = gdDBS.MoveRecords(dbcHogoshaMaster, -1) Then '//データは DESC ORDER かかっているのでこれでよい
        On Error GoTo spnRireki_SpinDownError
        '//金融機関の名称を表示
        Call SetBankAndShiten
'//最終のデータのみ編集可能とする
        If dbcHogoshaMaster.Recordset.IsFirst Then
            If eShoriKubun.Refer <> lblShoriKubun.Caption Then  '//参照以外の時
                dbcHogoshaMaster.Recordset.Edit     '//ここで排他が掛かる
                Call pLockedControl(False)
                spnRireki.Visible = True
                '//このボタンは支店をクリックした時に使えるようにする.
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
        Call MsgBox("これ以降にデータはありません.", vbInformation, mCaption)
    End If
    Exit Sub
spnRireki_SpinDownError:
    Call gdDBS.ErrorCheck   '//排他制御用エラートラップ
'    Call spnRireki_SpinUp
End Sub

Private Sub spnRireki_UpClick()
    '//前のレコードに移動
    If True = gdDBS.MoveRecords(dbcHogoshaMaster, 1) Then '//データは DESC ORDER かかっているのでこれでよい
        '//金融機関の名称を表示
        Call SetBankAndShiten
'//最終のデータのみ編集可能とする
'        dbcKeiyakushaMaster.Recordset.Edit
        Call mForm.LockedControlAllTextBox
        cmdEnd.Enabled = True
        cmdCancel.Enabled = True
    Else
        Call MsgBox("これ以前にデータはありません.", vbInformation, mCaption)
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
'//2007/06/07 カナ名と口座名義人名が同じ
'    '//現在の読みカナ名と口座名義人名が同じなら読みカナ名と口座名義人名に転送
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
    '// Return または Shift＋TAB のときのみ処理する
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
    'エラーの場合 KeyCode = 0 が返る
    If KeyCode = 0 Then
        Exit Sub
    End If
'//2006/04/26 前ゼロ埋め込み
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
        If eShoriKubun.Add <> lblShoriKubun.Caption Then     'レコード無しで新規以外の時
            msg = "該当データは存在しません.(" & lblHogoshaCode.Caption & ")"
        End If
    ElseIf eShoriKubun.Add = lblShoriKubun.Caption Then      'レコード有りで新規の時
        msg = "既にデータが存在します.(" & lblHogoshaCode.Caption & ")"
    End If
    If msg <> "" Then
        Call MsgBox(msg, vbInformation, mCaption)
        Call txtCAHGCD.SetFocus
        Exit Sub
    End If
    mIsActivated = False    '//レコード表示中のイベントを拾わないようにフラグを設定
    '//解約メッセージ抑止
    dbcHogoshaMaster.RecordSource = sql
    Call dbcHogoshaMaster.Refresh
    On Error GoTo txtCAHGCD_KeyDownError        '//排他制御用エラートラップ
    If 0& = dbcHogoshaMaster.Recordset.RecordCount Then
        '//新規登録
        dbcHogoshaMaster.Recordset.AddNew
        lblCAITKB.Caption = cboABKJNM.ItemData(cboABKJNM.ListIndex)
        lblCAKYCD.Caption = txtCAKYCD.Text
        lblCAKSCD.Caption = txtCAKSCD.Text
        lblCAHGCD.Caption = txtCAHGCD.Text
        lblCASQNO.Caption = gdDBS.sysDate("yyyymmdd")
        lblCAKKBN.Caption = 0
        lblCAKZSB.Caption = 1
        txtCAKYxx(0).Number = 20000101 '//一旦値を設定しないと「０」がセットされない：不思議？
        txtCAKYxx(0).Number = 0
        txtCAKYxx(1).Number = gdDBS.LastDay(0)
        txtCAFKxx(0).Number = 20000101 '//一旦値を設定しないと「０」がセットされない：不思議？
        txtCAFKxx(0).Number = 0
        txtCAFKxx(1).Number = gdDBS.LastDay(0)
    Else
        '//2007/06/06   銀行名・支店名の読み込みをここでするように変更
        '//             読込み時の Change()=名称表示 イベント順番が 支店コード・銀行コードの順になり支店名が表示されないことがある
        If eBankKubun.KinnyuuKikan = dbcHogoshaMaster.Recordset.Fields("CAKKBN").Value Then
            Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Bank, _
               dbcHogoshaMaster.Recordset.Fields("CABANK").Value, vDate:=gdDBS.sysDate("YYYYMMDD"))
            lblBankName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))
            Set dyn = Nothing
            Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Shiten, _
                dbcHogoshaMaster.Recordset.Fields("CABANK").Value, _
                dbcHogoshaMaster.Recordset.Fields("CASITN").Value, vDate:=gdDBS.sysDate("YYYYMMDD"))
            lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"支店名_漢字" で読めない
            Set dyn = Nothing
        End If
        '//修正・削除
        Call dbcHogoshaMaster.Recordset.MoveFirst
        Call dbcHogoshaMaster.Recordset.Edit
'        Call dbcHogoshaMaster.UpdateRecord
    End If
    '//参照で無ければボタンの制御開始
    If False = optShoriKubun(eShoriKubun.Refer).Value Then
        Call pLockedControl(False)
    End If
    spnRireki.Visible = dbcHogoshaMaster.Recordset.RecordCount > 1
    '//このボタンは支店をクリックした時に使えるようにする.
    cmdKakutei.Enabled = False
    '//解約メッセージ抑止
    '//コントロールを保護者（漢字）にしたいがためにおまじない：他に方法が見つからない？
    'If True = optShoriKubun(eShoriKubun.Refer).Value Then
        Call SendKeys("+{TAB}")
    'Else
    '    Call SendKeys("+{TAB}+{TAB}")
    'End If
    '//中止ボタンは参照以外はいつでも押下可能に！
    Call pButtonControl(optShoriKubun(eShoriKubun.Delete).Value, True)
    '//中止ボタンはいつでも押下可能に！
    If Not optShoriKubun(eShoriKubun.Refer).Value Then
        cmdCancel.Visible = True
        cmdCancel.Enabled = True
    End If
    Exit Sub
txtCAHGCD_KeyDownError:
    Call gdDBS.ErrorCheck(dbcHogoshaMaster.Database)    '//排他制御用エラートラップ
End Sub

Private Sub txtCAFKxx_LostFocus(Index As Integer)
    lblCAFKxx(Index).Caption = Val(gdDBS.Nz(txtCAFKxx(Index).Number))
End Sub

Private Sub txtCAKNNM_Change()
    txtCAKZNM.Text = txtCAKNNM.Text '//2007/06/07 保護者名(カナ)＝口座名義人名
    Call pButtonControl(True)
End Sub

Private Sub txtCAKSCD_LostFocus()
'//2006/04/26 前ゼロ埋め込み
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
    '// Return または Shift＋TAB のときのみ処理する
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
'//2006/04/26 前ゼロ埋め込み
    txtCAKYCD.Text = Format(Val(txtCAKYCD.Text), "00000")
'//2002/12/10 教室区分(??KSCD)は使用しない
'//    sql = "SELECT DISTINCT BAITKB,BAKYCD,BAKSCD,BAKJNM FROM tbKeiyakushaMaster"
    sql = "SELECT DISTINCT BAITKB,BAKYCD,BAKJNM FROM tbKeiyakushaMaster"
    sql = sql & " WHERE BAITKB = '" & cboABKJNM.ItemData(cboABKJNM.ListIndex) & "'"
    sql = sql & "   AND BAKYCD = '" & txtCAKYCD.Text & "'"
    sql = sql & "   AND TO_CHAR(SYSDATE,'YYYYMMDD') BETWEEN BAKYST AND BAKYED" '//有効データ絞込み
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If 0 = dyn.RecordCount Then
        Call dyn.Close
        KeyCode = 0
        '//                                        「契約者番号」
        Call MsgBox("契約者が解約状態、もしくは該当データが存在しません.(" & lblKeiyakushaCode.Caption & ")", vbInformation, mCaption)
        Call txtCAKYCD.SetFocus
        Exit Sub
    End If
    lblBAKJNM.Caption = dyn.Fields("BAKJNM")
#If 0 Then
'//2002/12/10 教室区分(??KSCD)は使用しない
    Call cboCAKSCDz.Clear
    Do Until dyn.EOF
'//2002/12/10 教室区分(??KSCD)は使用しない
'//        Call cboCAKSCDz.AddItem(dyn.Fields("BAKSCD"))
        Call dyn.MoveNext
    Loop
    cboCAKSCDz.ListIndex = 0
#End If
    Call dyn.Close
End Sub

Private Function pUpdateErrorCheck() As Boolean
'//2006/06/26 振込み依頼書にも同じロジックが有るので注意
    '///////////////////////////////
    '//必須入力項目と整合性チェック
    
    Dim str As New StringClass
    Dim obj As Object, msg As String
    '//保護者・漢字名称は必須
    If txtCAKJNM.Text = "" Then
        Set obj = txtCAKJNM
        msg = "保護者名(漢字)は必須入力です."
    ElseIf False = str.CheckLength(txtCAKJNM.Text) Then
        Set obj = txtCAKJNM
        msg = "保護者名(漢字)に半角が含まれています."
    End If
    '//保護者・カナ名称は必須
    '//2007/06/07 必須 復活：口座名義人と同じ値とする為
    If txtCAKNNM.Text = "" Then
        Set obj = txtCAKNNM
        msg = "保護者名(カナ)は必須入力です."
    ElseIf False = str.CheckLength(txtCAKNNM.Text, vbNarrow) Then
        Set obj = txtCAKNNM
        msg = "保護者名(カナ)に全角が含まれています."
    ElseIf 0 < InStr(txtCAKNNM.Text, "ｰ") Then
        Set obj = txtCAKNNM
        msg = "保護者名(カナ)に長音が含まれています."
    End If
    If IsNull(txtCAKYxx(1).Number) Then
        Set obj = txtCAKYxx(1)
        msg = "契約期間の終了日は必須入力です."
    ElseIf txtCAKYxx(0).Text > txtCAKYxx(1).Text Then
        Set obj = txtCAKYxx(0)
        msg = "契約期間が不正です."
    ElseIf IsNull(txtCAFKxx(1).Number) Then
        Set obj = txtCAFKxx(1)
        msg = "振替期間の終了日は必須入力です."
    ElseIf txtCAFKxx(0).Text > txtCAFKxx(1).Text Then
        Set obj = txtCAFKxx(0)
        msg = "振替期間が不正です."
    End If
    
    If lblCAKKBN.Caption = eBankKubun.KinnyuuKikan Then
        If txtCABANK.Text = "" Or lblBankName.Caption = "" Then
            Set obj = txtCABANK
            msg = "金融機関は必須入力です."
        ElseIf txtCASITN.Text = "" Or lblShitenName.Caption = "" Then
            Set obj = txtCASITN
            msg = "支店は必須入力です."
        ElseIf Not (lblCAKZSB.Caption = eBankYokinShubetsu.Futsuu _
                 Or lblCAKZSB.Caption = eBankYokinShubetsu.Touza) Then
            Set obj = optCAKZSB(eBankYokinShubetsu.Futsuu)
            msg = "預金種別は必須入力です."
        ElseIf txtCAKZNO.Text = "" Then
            Set obj = txtCAKZNO
            msg = "口座番号は必須入力です."
        End If
    ElseIf lblCAKKBN.Caption = eBankKubun.YuubinKyoku Then
        If txtCAYBTK.Text = "" Then
            Set obj = txtCAYBTK
            msg = "通帳記号は必須入力です."
        ElseIf txtCAYBTN.Text = "" Then
            Set obj = txtCAYBTN
            msg = "通帳番号は必須入力です."
        ElseIf "1" <> Right(txtCAYBTN.Text, 1) Then
'//2006/04/26 末尾番号チェック
            Set obj = txtCAYBTN
            msg = "通帳番号の末尾が「１」以外です."
        End If
    End If
    '//2007/06/07 必須 解除：口座名義人と同じ値とする為
'    If txtCAKZNM.Text = "" Then
'        Set obj = txtCAKZNM
'        msg = "口座名義人(カナ)は必須入力です."
'    End If
    '//Object が設定されているか？
    If TypeName(obj) <> "Nothing" Then
        Call MsgBox(msg, vbCritical, mCaption)
        Call obj.SetFocus
        Exit Function
    End If
    
    If lblCASQNO.Caption = gdDBS.sysDate("yyyymmdd") Then
        pUpdateErrorCheck = True    '//ＳＥＱが本日なのでそのまま更新
        Exit Function
    End If
    pUpdateErrorCheck = pRirekiAddNew()
    Exit Function
pUpdateErrorCheckError:
    Call gdDBS.ErrorCheck       '//エラートラップ
    pUpdateErrorCheck = False   '//安全のため：False で終了するはず
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
        Exit Function   '//新規登録なのでチェック無し
    End If
        
    If txtCAKJNM.Text <> gdDBS.Nz(dyn.Fields("CAKJNM")) _
    Or txtCAKZNM.Text <> gdDBS.Nz(dyn.Fields("CAKZNM")) Then
'''    If txtCAKJNM.Text <> gdDBS.Nz(dyn.Fields("CAKJNM")) _
'''    Or txtCAKNNM.Text <> gdDBS.Nz(dyn.Fields("CAKNNM")) Then
        AddRireki = "口座名義人"
    ElseIf lblCAKKBN.Caption <> gdDBS.Nz(dyn.Fields("CAKKBN")) Then
        AddRireki = "振替口座"
    ElseIf lblCAKKBN.Caption = eBankKubun.KinnyuuKikan Then
        '//金融機関情報が違えば履歴情報追加
        If txtCABANK.Text <> gdDBS.Nz(dyn.Fields("CABANK")) _
         Or txtCASITN.Text <> gdDBS.Nz(dyn.Fields("CASITN")) _
         Or lblCAKZSB.Caption <> gdDBS.Nz(dyn.Fields("CAKZSB")) _
         Or txtCAKZNO.Text <> gdDBS.Nz(dyn.Fields("CAKZNO")) Then
            AddRireki = "民間機関"
        End If
    ElseIf lblCAKKBN.Caption = eBankKubun.YuubinKyoku Then
        '//郵便局情報が違えば履歴情報追加
        If txtCAYBTK.Text <> gdDBS.Nz(dyn.Fields("CAYBTK")) _
         Or txtCAYBTN.Text <> gdDBS.Nz(dyn.Fields("CAYBTN")) Then
            AddRireki = "郵便局"
        End If
'''    ElseIf txtCAKZNM.Text <> gdDBS.Nz(dyn.Fields("CAKZNM")) Then
'''        AddRireki = "口座名義人"
    End If
    
    '///////////////////////////
    '//履歴作成しない場合終了
    If "" = AddRireki Then
        pRirekiAddNew = True    '//現在のレコードに更新
        Exit Function
    End If
    
    '///////////////////////////////////////////
    '//変更内容定義の画面を表示する
    Load frmMakeNewData
    With frmMakeNewData
        '//フォームをこのフォームの中央に位置付けする
        .Top = Me.Top + (Me.Height - .Height) / 2
        .Left = Me.Left + (Me.Width - .Width) / 2
        .lblMessage.Caption = "「" & AddRireki & "」の情報が変更されたため履歴を作成します." & vbCrLf & vbCrLf & _
                              "「追加」　履歴として過去の情報を残す場合はこのボタンを押します." & vbCrLf & _
                              "「上書き」現在のデータに上書きする場合はこのボタンを押します."
        .lblFurikomi.Caption = "振替開始日"
        Call .Show(vbModal)
        '//いつ破棄されるかわからないのでローカルコピーしておく
        Dim PushButton As Integer, KeiyakuEnd As Long, FurikaeEnd As Long
        PushButton = .mPushButton
        KeiyakuEnd = .mKeiyakuEnd
        FurikaeEnd = .mFurikaeEnd
        Set frmMakeNewData = Nothing
        If PushButton = ePushButton.Update Then
            pRirekiAddNew = True    '//現在のレコードに更新：この時だけ戻って更新する.
            Exit Function
        ElseIf PushButton = ePushButton.Cancel Then
            Exit Function
        End If
    End With
    '//ここから画面内容の更新及び履歴作成開始
    
    '//前もって追加するレコード削除
    sql = " DELETE tcHogoshaMaster"
    sql = sql & " WHERE CAITKB = '" & lblCAITKB.Caption & "'"
    sql = sql & "   AND CAKYCD = '" & lblCAKYCD.Caption & "'"
    sql = sql & "   AND CAKSCD = '" & lblCAKSCD.Caption & "'"
    sql = sql & "   AND CAHGCD = '" & lblCAHGCD.Caption & "'"
    sql = sql & "   AND CASQNO = -1"
    Call gdDBS.Database.ExecuteSQL(sql)
    
    '////////////////////////////////////////////////
    '//テーブル定義が変更された場合注意すること！！
    '//2007/06/11 遅かりし項目追加：あまり運用上していない？
    Dim FixedCol As String
    FixedCol = "CAITKB,CAKYCD,CAKSCD,CAKJNM,CAHGCD,CAKNNM," & _
               "CASTNM,CAKKBN,CABANK,CASITN,CAKZSB,CAKZNO," & _
               "CAKZNM,CAYBTK,CAYBTN,CAKYST,CAFKST,CASKGK," & _
               "CAHKGK,CAKYDT,CAKYFG,CATRFG,CAADDT,CAUSID," & _
               "CANWDT,CAKYSR,CACHEK"
                
    '現在の更新前データ退避
    sql = "INSERT INTO tcHogoshaMaster("
    sql = sql & "CASQNO,CAKYED,CAFKED,CAUPDT,"
    sql = sql & FixedCol
    sql = sql & ") SELECT "
    sql = sql & "-1,"
    '//入力された日の前月末日を設定
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
    
    '//画面の内容を更新:cmdUpdate()の一部関数を実行
    Call pUpdateRecord
    
    On Error GoTo pRirekiAddNewError
    '//画面のデータのＳＥＱを本日にする
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
    '//退避したデータのＳＥＱを変更前にする
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
'//2013/02/26 口座変更等の更新時の追加更新の際に２度 pUpdateRecord() が実行されるのを制御する
    mRirekiAddNewUpdate = True
    pRirekiAddNew = True
    Exit Function
pRirekiAddNewError:
    Call gdDBS.ErrorCheck       '//エラートラップ
    pRirekiAddNew = False   '//安全のため：False で終了するはず
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
    lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"支店名_漢字" で読めない
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
'//2006/04/26 前ゼロ埋め込み
    If "" <> txtCAYBTK.Text Then
        txtCAYBTK.Text = Format(Val(txtCAYBTK.Text), "000")
    End If
End Sub

Private Sub txtCAYBTN_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCAYBTN_LostFocus()
    '//2006/04/26 前ゼロ埋め込み
    If "" <> txtCAYBTN.Text Then
        If "1" <> Right(txtCAYBTN.Text, 1) Then
            Call MsgBox("末尾が「１」以外です.(" & lblTsuchoBango.Caption & ")", vbCritical, mCaption)
        Else
            txtCAYBTN.Text = Format(Val(txtCAYBTN.Text), "00000000")
        End If
    End If
End Sub
