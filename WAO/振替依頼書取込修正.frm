VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmFurikaeReqImportEdit 
   Caption         =   "振替依頼書(取込)修正"
   ClientHeight    =   7710
   ClientLeft      =   4455
   ClientTop       =   3060
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
   ScaleHeight     =   7710
   ScaleWidth      =   10125
   Begin VB.Frame fraSysDate 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   375
      Left            =   8640
      TabIndex        =   81
      Top             =   -60
      Width           =   1155
      Begin VB.Label lblSysDate 
         Caption         =   "Label1"
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
         TabIndex        =   82
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "戻る(&B)"
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
      Left            =   8220
      TabIndex        =   78
      Top             =   6945
      Width           =   1335
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
      Left            =   4140
      TabIndex        =   76
      Top             =   6945
      Width           =   1335
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
      Left            =   5640
      TabIndex        =   77
      Top             =   6945
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "前のデータ(&P)"
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
      Left            =   720
      TabIndex        =   74
      Top             =   6945
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "次のデータ(&N)"
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
      Left            =   2220
      TabIndex        =   75
      Top             =   6945
      Width           =   1335
   End
   Begin VB.ComboBox cboCIOKFG 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      ItemData        =   "振替依頼書取込修正.frx":0000
      Left            =   1800
      List            =   "振替依頼書取込修正.frx":000D
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   59
      Top             =   5010
      Width           =   2835
   End
   Begin VB.CheckBox chkCIMUPD 
      Caption         =   "マスタ反映しない"
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
      Left            =   2640
      TabIndex        =   57
      Top             =   4650
      Width           =   1935
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
      ItemData        =   "振替依頼書取込修正.frx":003D
      Left            =   1800
      List            =   "振替依頼書取込修正.frx":004A
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      TabStop         =   0   'False
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
      TabIndex        =   10
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
            Caption         =   "振替依頼書取込修正.frx":0068
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "振替依頼書取込修正.frx":00D4
            Key             =   "振替依頼書取込修正.frx":00F2
            MouseIcon       =   "振替依頼書取込修正.frx":0136
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
            Caption         =   "振替依頼書取込修正.frx":0152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "振替依頼書取込修正.frx":01BE
            Key             =   "振替依頼書取込修正.frx":01DC
            MouseIcon       =   "振替依頼書取込修正.frx":0220
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
            TabIndex        =   40
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
            TabIndex        =   39
            Top             =   480
            Width           =   1275
         End
      End
      Begin VB.OptionButton optCiKKBN 
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optCiKKBN 
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
         Caption         =   "振替依頼書取込修正.frx":023C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "振替依頼書取込修正.frx":02A8
         Key             =   "振替依頼書取込修正.frx":02C6
         MouseIcon       =   "振替依頼書取込修正.frx":030A
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
            Caption         =   "振替依頼書取込修正.frx":0326
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "振替依頼書取込修正.frx":0392
            Key             =   "振替依頼書取込修正.frx":03B0
            MouseIcon       =   "振替依頼書取込修正.frx":03F4
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
            Caption         =   "振替依頼書取込修正.frx":0410
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "振替依頼書取込修正.frx":047C
            Key             =   "振替依頼書取込修正.frx":049A
            MouseIcon       =   "振替依頼書取込修正.frx":04DE
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
            Caption         =   "振替依頼書取込修正.frx":04FA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "振替依頼書取込修正.frx":0566
            Key             =   "振替依頼書取込修正.frx":0584
            MouseIcon       =   "振替依頼書取込修正.frx":05C8
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
            TabIndex        =   41
            Top             =   900
            Width           =   2535
            Begin VB.OptionButton optCiKZSB 
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
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   180
               Width           =   675
            End
            Begin VB.OptionButton optCiKZSB 
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
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   180
               Width           =   675
            End
            Begin VB.OptionButton optCiKZSB 
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
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   480
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label lblCiKZSB 
               BackColor       =   &H000000FF&
               Caption         =   "口座種別"
               DataField       =   "CiKZSB"
               DataSource      =   "dbcImportEdit"
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
               TabIndex        =   42
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
            Top             =   660
            Width           =   1935
         End
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
         TabIndex        =   55
         Top             =   2340
         Width           =   1395
      End
      Begin VB.Label lblCiKKBN 
         BackColor       =   &H000000FF&
         Caption         =   "金融機関種別"
         DataField       =   "CiKKBN"
         DataSource      =   "dbcImportEdit"
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
         TabIndex        =   49
         Top             =   180
         Width           =   1095
      End
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
      TabIndex        =   26
      Top             =   3360
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
         ItemData        =   "振替依頼書取込修正.frx":05E4
         Left            =   1500
         List            =   "振替依頼書取込修正.frx":0609
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   27
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
         ItemData        =   "振替依頼書取込修正.frx":064B
         Left            =   3900
         List            =   "振替依頼書取込修正.frx":0670
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   29
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
         Bindings        =   "振替依頼書取込修正.frx":06B2
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
         Bindings        =   "振替依頼書取込修正.frx":06C8
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
         TabIndex        =   51
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
         Name            =   "ＭＳ Ｐゴシック"
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
      Calendar        =   "振替依頼書取込修正.frx":06E0
      Caption         =   "振替依頼書取込修正.frx":0860
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":08CC
      Keys            =   "振替依頼書取込修正.frx":08EA
      MouseIcon       =   "振替依頼書取込修正.frx":0948
      Spin            =   "振替依頼書取込修正.frx":0964
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
      Caption         =   "振替依頼書取込修正.frx":098C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":09F8
      Key             =   "振替依頼書取込修正.frx":0A16
      MouseIcon       =   "振替依頼書取込修正.frx":0A5A
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
      Caption         =   "振替依頼書取込修正.frx":0A76
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":0AE2
      Key             =   "振替依頼書取込修正.frx":0B00
      MouseIcon       =   "振替依頼書取込修正.frx":0B44
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
      Caption         =   "振替依頼書取込修正.frx":0B60
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":0BCC
      Key             =   "振替依頼書取込修正.frx":0BEA
      MouseIcon       =   "振替依頼書取込修正.frx":0C2E
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
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "振替依頼書取込修正.frx":0C4A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":0CB6
      Key             =   "振替依頼書取込修正.frx":0CD4
      MouseIcon       =   "振替依頼書取込修正.frx":0D18
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
      Caption         =   "振替依頼書取込修正.frx":0D34
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":0DA0
      Key             =   "振替依頼書取込修正.frx":0DBE
      MouseIcon       =   "振替依頼書取込修正.frx":0E02
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
      Text            =   "ｶﾅｼﾒｲ........................*"
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
      Calendar        =   "振替依頼書取込修正.frx":0E1E
      Caption         =   "振替依頼書取込修正.frx":0F9E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":100A
      Keys            =   "振替依頼書取込修正.frx":1028
      MouseIcon       =   "振替依頼書取込修正.frx":1086
      Spin            =   "振替依頼書取込修正.frx":10A2
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
      Caption         =   "振替依頼書取込修正.frx":10CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":1126
      Key             =   "振替依頼書取込修正.frx":1144
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
      Text            =   "警告メッセージが複数行に表示される。"
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
      Caption         =   "振替依頼書取込修正.frx":1188
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":11F4
      Key             =   "振替依頼書取込修正.frx":1212
      MouseIcon       =   "振替依頼書取込修正.frx":1256
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
      Text            =   "金融機関名．．．．．．．．．＊"
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
      Caption         =   "振替依頼書取込修正.frx":1272
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":12DE
      Key             =   "振替依頼書取込修正.frx":12FC
      MouseIcon       =   "振替依頼書取込修正.frx":1340
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
      Text            =   "支店名．．．．．．．．．．．＊"
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
         Caption         =   "置換え"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   85
         Top             =   150
         Width           =   900
      End
      Begin VB.OptionButton optCiINSD 
         Caption         =   "追加"
         Height          =   240
         Index           =   1
         Left            =   1200
         TabIndex        =   84
         Top             =   150
         Width           =   690
      End
      Begin VB.Label lblCiINSD 
         BackColor       =   &H000000FF&
         Caption         =   "更新方法"
         DataField       =   "CiINSD"
         DataSource      =   "dbcImportEdit"
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
         Left            =   0
         TabIndex        =   87
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Label lblUpdMode 
      Alignment       =   1  '右揃え
      Caption         =   "更新方法"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "支店名(取込)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "金融機関名(取込)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "更新日"
      DataField       =   "CIUPDT"
      DataSource      =   "dbcImportEdit"
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
      Left            =   3300
      TabIndex        =   73
      Top             =   7995
      Width           =   915
   End
   Begin VB.Label lblCIUSID 
      BackColor       =   &H000000FF&
      Caption         =   "更新者"
      DataField       =   "CIUSID"
      DataSource      =   "dbcImportEdit"
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
      Left            =   2280
      TabIndex        =   72
      Top             =   7995
      Width           =   795
   End
   Begin VB.Label lblCIWMSG 
      BackColor       =   &H000000FF&
      Caption         =   "警告メッセージが複数行に表示される。"
      DataField       =   "CIWMSG"
      DataSource      =   "dbcImportEdit"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Caption         =   "変更後-F"
      DataField       =   "CIERROR"
      DataSource      =   "dbcImportEdit"
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
      Left            =   1560
      TabIndex        =   70
      Top             =   7635
      Width           =   795
   End
   Begin VB.Label lblCIERSR 
      BackColor       =   &H000000FF&
      Caption         =   "変更前-F"
      DataField       =   "CIERSR"
      DataSource      =   "dbcImportEdit"
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
      Left            =   720
      TabIndex        =   69
      Top             =   7635
      Width           =   795
   End
   Begin VB.Label lblCIOKFG 
      BackColor       =   &H000000FF&
      Caption         =   "反映ＯＫ"
      DataField       =   "CIOKFG"
      DataSource      =   "dbcImportEdit"
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
      Left            =   2400
      TabIndex        =   68
      Top             =   7635
      Width           =   855
   End
   Begin VB.Label lblCIMUPD 
      BackColor       =   &H000000FF&
      Caption         =   "反映しない"
      DataField       =   "CIMUPD"
      DataSource      =   "dbcImportEdit"
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
      Left            =   3360
      TabIndex        =   67
      Top             =   7635
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "取込日時-SEQ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      BorderStyle     =   1  '実線
      Caption         =   "取込SEQ"
      DataField       =   "CISEQN"
      DataSource      =   "dbcImportEdit"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "－"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "2006/03/01 23:59:59"
      DataField       =   "CIINDT"
      DataSource      =   "dbcImportEdit"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Left            =   240
      Picture         =   "振替依頼書取込修正.frx":135C
      Top             =   5370
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label10x 
      Alignment       =   1  '右揃え
      Caption         =   "マスタ反映方法"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      Caption         =   "異常"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Alignment       =   1  '右揃え
      Caption         =   "処理結果"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      TabIndex        =   56
      Top             =   2355
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
      TabIndex        =   54
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label lblCiITKB 
      BackColor       =   &H000000FF&
      Caption         =   "委託者区分"
      DataField       =   "CiITKB"
      DataSource      =   "dbcImportEdit"
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
      Left            =   3660
      TabIndex        =   53
      Top             =   840
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
      Height          =   195
      Left            =   360
      TabIndex        =   52
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label lblCiSQNO 
      BackColor       =   &H000000FF&
      Caption         =   "保護者ＳＥＱ"
      DataField       =   "CiSQNO"
      DataSource      =   "dbcImportEdit"
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
      Left            =   3960
      TabIndex        =   14
      Top             =   1725
      Width           =   975
   End
   Begin VB.Label lblCiKYCD 
      BackColor       =   &H000000FF&
      Caption         =   "オーナー番号"
      DataField       =   "CiKYCD"
      DataSource      =   "dbcImportEdit"
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
      Left            =   3660
      TabIndex        =   38
      Top             =   1140
      Width           =   1155
   End
   Begin VB.Label lblCiHGCD 
      BackColor       =   &H000000FF&
      Caption         =   "保護者番号"
      DataField       =   "CiHGCD"
      DataSource      =   "dbcImportEdit"
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
      Left            =   2850
      TabIndex        =   13
      Top             =   1725
      Width           =   975
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
      Left            =   2760
      TabIndex        =   37
      Top             =   1380
      Width           =   2235
   End
   Begin VB.Label Label10 
      Alignment       =   1  '右揃え
      BackColor       =   &H000000FF&
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
      TabIndex        =   36
      Top             =   3315
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblKeiyakushaCode 
      Alignment       =   1  '右揃え
      Caption         =   "オーナー番号"
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
      Left            =   360
      TabIndex        =   35
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label lblHogoshaCode 
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
      Height          =   195
      Left            =   480
      TabIndex        =   34
      Top             =   1695
      Width           =   1155
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
      Left            =   285
      TabIndex        =   33
      Top             =   2040
      Width           =   1395
   End
   Begin VB.Label Label18 
      Alignment       =   1  '右揃え
      Caption         =   "振替開始年月"
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
      TabIndex        =   32
      Top             =   3315
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
Attribute VB_Name = "frmFurikaeReqImportEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm As New FormClass
Private mCaption As String
Private mBankChange As Boolean  '//2006/08/22 ???_Change イベントを銀行=>支店に強制する

Private mErrMsgOn As Boolean
Private mCheckUpdate As Boolean
Private mRimp As New FurikaeReqImpClass
Private mUpdateOK As Boolean
Private mIsActivated As Boolean

'//2007/06/07 更新・中止ボタンを完全単独にコントロール
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
    
    Call mForm.LockedControl(False) '//常にデータは編集可能にしておく
'    cmdUpdate.Enabled = blMode
'    cmdCancel.Enabled = blMode
    'mForm.LockedControl() で警告表示が赤色の為、消える！
    lblERRMSG.Visible = True
    '//2007/06/07 口座名義人は常に入力しない：保護者名(カナ)をコピーする様に仕様変更
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
        '//キーを押した時に更新可能か判断
'        cmdUpdate.Enabled = mCheckUpdate    '//更新ボタンの制御：データ表示時にイベントが発生しても可能なように！
'        cmdCancel.Enabled = cmdUpdate.Enabled
    End If
    Call pButtonControl(True)
End Sub

Private Sub cboCIOKFG_Click()
    '//修正前のエラーにより選択内容を制御する
    Select Case lblCIERSR.Caption
    Case mRimp.errEditData    '//ありえない
    Case mRimp.errInvalid, mRimp.errImport
        If cboCIOKFG.ItemData(cboCIOKFG.ListIndex) <> mRimp.updInvalid Then
            '//一切の選択不可能！！！
            Call MsgBox("「取込」又は「異常」データの為、選択できません。" & vbCrLf & "チェック処理を実行して下さい。", vbCritical + vbOKOnly, mCaption)
            '//cboCIOKFG.ListIndex = mRimp.updInvalid + 2     '// -2 ～ 2
            '//元に戻す
            cboCIOKFG.ListIndex = Val(lblCIOKFG.Caption) + 2    '// -2 ～ 2
            Exit Sub
        End If
    Case mRimp.errNormal
        '//何でもＯＫ
        '//2014/06/11 解約状態で無いのに解約解除を選択した場合
        If False = checkKaiyaku() Then
            If cboCIOKFG.ItemData(cboCIOKFG.ListIndex) = mRimp.updResetCancel Then
                '//解約解除は関係ない
                Call MsgBox("解約状態ではありません。", vbInformation + vbOKOnly, mCaption)
                '//元に戻す
                cboCIOKFG.ListIndex = Val(lblCIOKFG.Caption) + 2    '// -2 ～ 2
            End If
        End If
    Case mRimp.errWarning
        If cboCIOKFG.ItemData(cboCIOKFG.ListIndex) = mRimp.updNormal Then
            '//再チェック時に警告に戻るので選択の意味が無い
            Call MsgBox("「警告」データを反映するには" & vbCrLf & "「" & mRimp.mUpdateMessage(mRimp.updWarnUpd) & "」を選択してください。", vbInformation + vbOKOnly, mCaption)
            '//元に戻す
            cboCIOKFG.ListIndex = Val(lblCIOKFG.Caption) + 2    '// -2 ～ 2
            Exit Sub
        Else
            If checkKaiyaku() Then
            'If InStr(lblCIWMSG.Caption, "解約状態") Then
                If cboCIOKFG.ItemData(cboCIOKFG.ListIndex) = mRimp.updWarnUpd Then
                    '//解約解除しなくて良いか
                    If vbOK <> MsgBox("解約状態は解除されません。" & vbCrLf & "よろしいですか？", vbInformation + vbOKCancel, mCaption) Then
                        Exit Sub
                    End If
                End If
            ElseIf cboCIOKFG.ItemData(cboCIOKFG.ListIndex) = mRimp.updResetCancel Then
                '//解約解除は関係ない
                Call MsgBox("解約状態ではありません。", vbInformation + vbOKOnly, mCaption)
                '//元に戻す
                cboCIOKFG.ListIndex = Val(lblCIOKFG.Caption) + 2    '// -2 ～ 2
            End If
        End If
    Case Else                   '//ありえない
    End Select
    lblCIOKFG.Caption = cboCIOKFG.ItemData(cboCIOKFG.ListIndex)
    '//2014/06/09 コンボボックス変更時にボタンを使用可能に
    Call pButtonControl(True)
    '//キーを押した時に更新可能か判断
'    cmdUpdate.Enabled = mCheckUpdate    '//更新ボタンの制御：データ表示時にイベントが発生しても可能なように！
'    cmdCancel.Enabled = cmdUpdate.Enabled
    'Call SendKeys("{TAB}")  '//結果を正しく見せたいのでフォーカス移動
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
    Call sttausCIERROR  '//2014/05/19 データがイベント毎にいろいろ発生するのでここに統一
End Sub

Private Sub cmdCancel_Click()
    Call dbcImportEdit.UpdateControls
    'Call cboABKJNM.SetFocus
    Call pLockedControl(False)
    Call lblCIERROR_Change
    Call pButtonControl(False)
    Call sttausCIERROR  '//2014/05/19 データがイベント毎にいろいろ発生するのでここに統一
End Sub

Private Function pCheckEditData() As Boolean
    Dim obj As Object, Edit As Boolean
    For Each obj In Me.Controls
        If TypeOf obj Is imText _
        Or TypeOf obj Is imNumber _
        Or TypeOf obj Is imDate _
        Or TypeOf obj Is Label Then
            '//コントロールの DataChanged プロパティを検査して更新を必要とするか判断
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
    '//入力内容チェックで取りやめしたので終了
    mUpdateOK = pUpdateErrorCheck
    If False = mUpdateOK Then
        Exit Sub
    End If
    mUpdateOK = True
    lblCIERROR.Caption = mRimp.errEditData    '//編集後は必ずエラーフラグを立てる：チェック処理を必ずする
    lblCIUSID.Caption = gdDBS.LoginUserName
    lblCIUPDT.Caption = gdDBS.sysDate
    '//メインの SpreadSheet に内容を反映する：Update後では DataChanged() が変化してしまうので...。
    Call frmFurikaeReqImport.gEditToSpreadSheet(0)
    '//画面の内容をＤＢに更新
    Call dbcImportEdit.UpdateRecord
    'Call pErrorCheck
    Call pLockedControl(False)
    Call lblCIERROR_Change
    Call pButtonControl(False)
    Call sttausCIERROR  '//2014/05/19 データがイベント毎にいろいろ発生するのでここに統一
End Sub

Public Sub cmdEnd_Click()
    If True = pCheckEditData Then
        Dim stts As Integer
        stts = MsgBox("内容が変更されています。" & vbCrLf & vbCrLf & "更新しますか？", vbYesNoCancel + vbInformation, mCaption)
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
    Call frmFurikaeReqImport.Show  '//強制的に飛び元の画面を表示
    Unload Me
End Sub

Private Sub cmdNext_Click()
    mIsActivated = False
    If True = pCheckEditData Then
        Dim stts As Integer
        stts = MsgBox("内容が変更されています。" & vbCrLf & vbCrLf & "更新しますか？", vbYesNoCancel + vbInformation, mCaption)
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
    '//メインの SpreadSheet に内容を反映する：Update後では DataChanged() が変化してしまうので...。
    frmFurikaeReqImport.mEditRow = frmFurikaeReqImport.mEditRow + 1
    '//これから編集するのに既に編集済みとなっているのを回避する
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
        stts = MsgBox("内容が変更されています。" & vbCrLf & vbCrLf & "更新しますか？", vbYesNoCancel + vbInformation, mCaption)
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
    '//メインの SpreadSheet に内容を反映する：Update後では DataChanged() が変化してしまうので...。
    frmFurikaeReqImport.mEditRow = frmFurikaeReqImport.mEditRow - 1
    '//これから編集するのに既に編集済みとなっているのを回避する
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
    '//入力された金融機関名＆支店名を強制書き換え
    txtCiBKNM.Text = Mid(dblBankList.Text, 6)
    lblBankName.Caption = Mid(dblBankList.Text, 6)
    txtCiSINM.Text = Mid(dblShitenList.Text, 5)
    lblShitenName.Caption = Mid(dblShitenList.Text, 5)
    cmdKakutei.Enabled = False
'//2006/08/22 確定後交信可能に！
    Call pLockedControl(True)
End Sub

'///////////////////////////////////////////////////////
'//レコード移動時にこのイベントが起きる：編集を開始
Private Sub dbcImportEdit_Reposition()
    cmdNext.Enabled = Not dbcImportEdit.Recordset.IsLast
    cmdPrev.Enabled = Not dbcImportEdit.Recordset.IsFirst
    If dbcImportEdit.Recordset.BOF _
    Or dbcImportEdit.Recordset.EOF Then
        '//先頭以前、最後以降のレコード位置は編集開始をしない
        Exit Sub
    End If
    'Debug.Print dbcImportEdit.Recordset.RowPosition
    '//各入力項目のエラー表示
    Dim obj As Object
    For Each obj In Controls
        If TypeOf obj Is imText _
        Or TypeOf obj Is imNumber _
        Or TypeOf obj Is imDate Then
            If "" <> obj.DataField Then
                '//全項目 ORADC にバインドされているはず！
                obj.BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(obj.DataField & "E"))
            End If
        End If
    Next obj
    '//委託者コードのエラー表示
    cboABKJNM.BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCiITKB.DataField & "E"))
    '//金融機関区分のエラー表示
    optCiKKBN(0).BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCiKKBN.DataField & "E"), False)
    optCiKKBN(1).BackColor = optCiKKBN(0).BackColor
    '//預金種別のエラー表示
    optCiKZSB(0).BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCiKZSB.DataField & "E"), False)
    optCiKZSB(1).BackColor = optCiKZSB(0).BackColor
    optCiKZSB(2).BackColor = optCiKZSB(0).BackColor
    cboCIOKFG.ListIndex = Val(lblCIOKFG.Caption) + 2    '// -2 ～ 2
    chkCIMUPD.Value = Abs(Val(lblCIMUPD.Caption) <> 0)
    
    Call sttausCIERROR  '//2014/05/19 データがイベント毎にいろいろ発生するのでここに統一
    
    Call dbcImportEdit.Recordset.Edit         '//編集開始
End Sub

Private Sub dblBankList_Click()
    cboShitenYomi.ListIndex = -1
    Call cboShitenYomi_Click
End Sub

Private Sub dblShitenList_Click()
    cmdKakutei.Enabled = dblBankList.Text <> ""
End Sub

Private Sub Form_Activate()
    mCheckUpdate = True     '//更新ボタンの制御：データ表示時にイベントが発生しても可能なように！
    If False = mIsActivated Then
        Call pButtonControl(False, True)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
    mErrMsgOn = True
    '//キーを押した時に更新可能か判断
'    cmdUpdate.Enabled = pCheckEditData
'    cmdCancel.Enabled = cmdUpdate.Enabled
End Sub

Private Sub Form_Load()
    mCheckUpdate = False    '//更新ボタンの制御：データ表示時にイベントが発生しても可能なように！
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    Call mForm.MoveSysDate
    '//銀行と郵便局の Frame を整列する
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
    '//呼び出し元で設定するので不要
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
    '//これ以上小さくするとコントロールが隠れるので制御する
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
    '//子フォームとして存在するのを破棄
    Set gdFormSub = Nothing
    Set mForm = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

'//2014/05/19 データがイベント毎にいろいろ発生するのでここに統一
Private Sub sttausCIERROR()
    Dim err As Integer
    err = Val(lblCIERROR.Caption)
    If err = mRimp.errInvalid And 0 <> Val(lblCIMUPD.Caption) Then
        err = mRimp.errWarning
    End If
    Select Case err
    Case mRimp.errImport:     lblERRMSG.Caption = "取込": lblERRMSG.BackColor = mRimp.ErrorStatus(err)
    Case mRimp.errEditData:   lblERRMSG.Caption = "修正": lblERRMSG.BackColor = mRimp.ErrorStatus(err)
    Case mRimp.errInvalid:    lblERRMSG.Caption = "異常": lblERRMSG.BackColor = mRimp.ErrorStatus(err)
    Case mRimp.errNormal:     lblERRMSG.Caption = "正常": lblERRMSG.BackColor = vbCyan
    Case mRimp.errWarning:    lblERRMSG.Caption = "警告": lblERRMSG.BackColor = mRimp.ErrorStatus(err)
    Case Else:                lblERRMSG.Caption = "例外": lblERRMSG.BackColor = vbRed
    End Select
    'lblERRMSG.BackColor = mRimp.ErrorStatus(lblCIERROR.Caption)
    '//2014/05/19 更新モードの追加
    If err = mRimp.errInvalid Then
        '//異常データ時には使用できないように制御する
        fraCiINSD.Enabled = False
    Else
        '//保護者マスタにデータが無い時には使用できないように制御する
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
    '//ブランクはエラーとする
    If Not IsNull(lblCiKKBN.Caption) And "" <> lblCiKKBN.Caption Then
        optCiKKBN(lblCiKKBN.Caption).Value = True
    End If
End Sub

Private Sub lblCIKZSB_Change()
    If Not IsNull(lblCiKZSB.Caption) And "" <> lblCiKZSB.Caption Then
        optCiKZSB(Val(lblCiKZSB.Caption)).Value = True
    Else
'//設定すると更新フラグが立ってしまうので止める
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
    '//フォーカスが消えるので設定する.
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
    '//ブランクはエラーとする
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
        
'//2013/06/18 前ゼロ埋め込み
    txtCiKYCD.Text = Format(Val(txtCiKYCD.Text), String(7, "0"))
''    If "" = Trim(txtCiKYCD.Text) Then
''        Exit Sub
''    End If
'//2002/12/10 教室区分(??KSCD)は使用しない
'//    sql = "SELECT DISTINCT BAITKB,BAKYCD,BAKSCD,BAKJNM FROM tbKeiyakushaMaster"
'//2015/02/12 最新データ１件のみで判別
'//    sql = "SELECT DISTINCT BAITKB,BAKYCD,BAKJNM,BAKYED FROM tbKeiyakushaMaster"
    sql = "SELECT BAITKB,BAKYCD,BAKJNM,BAKYED FROM tbKeiyakushaMaster"
    sql = sql & " WHERE BAITKB = '" & lblCiITKB.Caption & "'"
    sql = sql & "   AND BAKYCD = '" & txtCiKYCD.Text & "'"
'//2006/03/31 解約状態を表示するように変更
'    sql = sql & "   AND TO_CHAR(SYSDATE,'YYYYMMDD') BETWEEN BAKYST AND BAKYED" '//有効データ絞込み
'//2015/02/12 最新データ１件のみで判別
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
            '//                                        「契約者番号」
            Call MsgBox("該当データは存在しません.(" & lblKeiyakushaCode.Caption & ")", vbInformation, mCaption)
            Call txtCiKYCD.SetFocus
        End If
        'Exit Sub
    Else
#If 1 Then
'//2015/02/12 解約状態を表示するように変更
        lblBAKJNM.ForeColor = vbBlack
        Dim kk As String
        If dyn.Fields("BAKYED") < gdDBS.sysDate("yyyymmdd") Then
            kk = "(解約)"
            lblBAKJNM.ForeColor = vbRed
        End If
        lblBAKJNM.Caption = kk & dyn.Fields("BAKJNM")
#Else
        lblBAKJNM.Caption = IIf(dyn.Fields("BAKYED") < gdDBS.sysDate("yyyymmdd"), "(解約)", "") & _
                            dyn.Fields("BAKJNM")
#End If
    End If

#If 0 Then
'//2002/12/10 教室区分(??KSCD)は使用しない
    Call cboCIKSCDz.Clear
    Do Until dyn.EOF
'//2002/12/10 教室区分(??KSCD)は使用しない
'//        Call cboCIKSCDz.AddItem(dyn.Fields("BAKSCD"))
        Call dyn.MoveNext
    Loop
    cboCIKSCDz.ListIndex = 0
#End If
    Call dyn.Close
    '//2007/06/06   銀行名・支店名の読み込みをここでするように変更
    '//             読込み時の Change()=名称表示 イベント順番が 支店コード・銀行コードの順になり支店名が表示されないことがある
    Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Bank, txtCiBANK.Text, vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblBankName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))
    Set dyn = Nothing
    Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Shiten, txtCiBANK.Text, txtCiSITN.Text, vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"支店名_漢字" で読めない
    Set dyn = Nothing
    'txtCIKJNM.SetFocus
End Sub

Private Function pUpdateErrorCheck() As Boolean
'//2012/07/11 マスタ反映しない場合チェックしない
    If chkCIMUPD.Value <> 0 Then
        pUpdateErrorCheck = True
        Exit Function
    End If
'//2006/06/26 更新時のチェックがなかったので追加：保護者メンテをコピー
    '///////////////////////////////
    '//必須入力項目と整合性チェック
    
    Dim str As New StringClass
    Dim obj As Object, msg As String
    '//保護者・漢字名称は必須
    If txtCiKJNM.Text = "" Then
        Set obj = txtCiKJNM
        msg = "保護者名(漢字)は必須入力です."
    ElseIf False = str.CheckLength(txtCiKJNM.Text) Then
        Set obj = txtCiKJNM
        msg = "保護者名(漢字)に半角が含まれています."
    End If
    '//保護者・カナ名称は必須
    '//2007/06/07 必須 復活：口座名義人と同じ値とする為
    If txtCiKNNM.Text = "" Then
        Set obj = txtCiKNNM
        msg = "保護者名(カナ)は必須入力です."
    ElseIf False = str.CheckLength(txtCiKNNM.Text, vbNarrow) Then
        Set obj = txtCiKNNM
        msg = "保護者名(カナ)に全角が含まれています."
    ElseIf 0 < InStr(txtCiKNNM.Text, "ｰ") Then
        Set obj = txtCiKNNM
        msg = "保護者名(カナ)に長音が含まれています."
    End If
#If 0 Then  '//項目なし
    If IsNull(txtCIKYxx(1).Number) Then
        Set obj = txtCIKYxx(1)
        msg = "契約期間の終了日は必須入力です."
    ElseIf txtCIKYxx(0).Text > txtCIKYxx(1).Text Then
        Set obj = txtCIKYxx(0)
        msg = "契約期間が不正です."
    ElseIf IsNull(txtCiFKxx(1).Number) Then
        Set obj = txtCiFKxx(1)
        msg = "振替期間の終了日は必須入力です."
    ElseIf txtCiFKxx(0).Text > txtCiFKxx(1).Text Then
        Set obj = txtCiFKxx(0)
        msg = "振替期間が不正です."
    End If
#End If
    If lblCiKKBN.Caption = "" Then
        If txtCiBANK.Visible = True And txtCiBANK.Enabled = True Then
            Set obj = txtCiBANK
        ElseIf txtCiYBTK.Visible = True And txtCiYBTK.Enabled = True Then
            Set obj = txtCiYBTK
        Else
            Set obj = txtCiKYCD     '// ==> fraKinnyuuKikan にはフォーカスを当てられないのでここを強制
        End If
        msg = "金融機関区分は必須入力です."
    ElseIf lblCiKKBN.Caption = eBankKubun.KinnyuuKikan Then
        If txtCiBANK.Text = "" Or lblBankName.Caption = "" Then
            Set obj = txtCiBANK
            msg = "金融機関は必須入力です."
        ElseIf txtCiSITN.Text = "" Or lblShitenName.Caption = "" Then
            Set obj = txtCiSITN
            msg = "支店は必須入力です."
        ElseIf Not (lblCiKZSB.Caption = eBankYokinShubetsu.Futsuu _
                 Or lblCiKZSB.Caption = eBankYokinShubetsu.Touza) Then
            Set obj = optCiKZSB(eBankYokinShubetsu.Futsuu)
            msg = "預金種別は必須入力です."
        ElseIf txtCiKZNO.Text = "" Then
            Set obj = txtCiKZNO
            msg = "口座番号は必須入力です."
        End If
    ElseIf lblCiKKBN.Caption = eBankKubun.YuubinKyoku Then
        If txtCiYBTK.Text = "" Then
            Set obj = txtCiYBTK
            msg = "通帳記号は必須入力です."
        ElseIf txtCiYBTN.Text = "" Then
            Set obj = txtCiYBTN
            msg = "通帳番号は必須入力です."
        ElseIf "1" <> Right(txtCiYBTN.Text, 1) Then
'//2006/04/26 末尾番号チェック
            Set obj = txtCiYBTN
            msg = "通帳番号の末尾が「１」以外です."
        End If
    End If
    If txtCiKZNM.Text = "" Then
        Set obj = txtCiKZNM
        msg = "口座名義人(カナ)は必須入力です."
    End If
    '//Object が設定されているか？
    If TypeName(obj) <> "Nothing" Then
        Call MsgBox(msg, vbCritical, mCaption)
        Call obj.SetFocus
        Exit Function
    End If
    pUpdateErrorCheck = True
    Exit Function
pUpdateErrorCheckError:
    Call gdDBS.ErrorCheck       '//エラートラップ
    pUpdateErrorCheck = False   '//安全のため：False で終了するはず
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
    lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"支店名_漢字" で読めない
End Sub

Private Sub txtCIYBTK_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCIYBTN_Change()
    Call pButtonControl(True)
End Sub

'/////////////////////////
'//再エラーチェック再考！！！レコードの更新が出来なくなる
Private Function pErrorCheck()
    '//各入力項目のエラー表示
    Dim obj As Object
    
    Call frmFurikaeReqImport.gDataCheck(Format(lblCIINDT.Caption, "yyyy/MM/dd hh:nn:ss"), lblCISEQN.Caption)
    For Each obj In Controls
        If TypeOf obj Is imText _
        Or TypeOf obj Is imNumber _
        Or TypeOf obj Is imDate Then
            If "" <> obj.DataField Then
                '//全項目 ORADC にバインドされているはず！
                obj.BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(obj.DataField & "E"))
            End If
        End If
    Next obj
    '//委託者コードのエラー表示
    cboABKJNM.BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCiITKB.DataField & "E"))
    '//金融機関区分のエラー表示
    optCiKKBN(0).BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCiKKBN.DataField & "E"), False)
    optCiKKBN(1).BackColor = optCiKKBN(0).BackColor
    '//預金種別のエラー表示
    optCiKZSB(0).BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCiKZSB.DataField & "E"), False)
    optCiKZSB(1).BackColor = optCiKZSB(0).BackColor
    optCiKZSB(2).BackColor = optCiKZSB(0).BackColor
End Function

'//保護者マスタのレコードが既に存在するか
Private Function checkExists()
    checkExists = InStr(lblCIWMSG.Caption, MainModule.cEXISTS_DATA) <> 0 _
               Or InStr(lblCIWMSG.Caption, MainModule.cKAIYAKU_DATA) <> 0
End Function

'//保護者マスタのレコードが存在し、解約状態であるか
Private Function checkKaiyaku()
    checkKaiyaku = InStr(lblCIWMSG.Caption, MainModule.cKAIYAKU_DATA) <> 0
End Function

