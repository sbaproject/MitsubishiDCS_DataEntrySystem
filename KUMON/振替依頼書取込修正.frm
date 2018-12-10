VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{7CC4CE40-1297-11D2-9BBF-00A024695830}#1.0#0"; "Number60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmFurikaeReqImportEdit 
   Caption         =   "振替依頼書(取込)修正"
   ClientHeight    =   8040
   ClientLeft      =   1950
   ClientTop       =   2340
   ClientWidth     =   10470
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
   ScaleHeight     =   8040
   ScaleWidth      =   10470
   Begin VB.CheckBox chkCIMUPD 
      Caption         =   "マスタ反映しない"
      Height          =   255
      Left            =   2640
      TabIndex        =   78
      Top             =   4140
      Width           =   1935
   End
   Begin imText6Ctl.imText txtCIWMSG 
      Height          =   1575
      Left            =   780
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4860
      Width           =   4155
      _Version        =   65536
      _ExtentX        =   7329
      _ExtentY        =   2778
      Caption         =   "振替依頼書取込修正.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":005C
      Key             =   "振替依頼書取込修正.frx":007A
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
   Begin VB.Frame fraSysDate 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   375
      Left            =   9060
      TabIndex        =   73
      Top             =   0
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
         TabIndex        =   74
         Top             =   60
         Width           =   855
      End
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
      ItemData        =   "振替依頼書取込修正.frx":00BE
      Left            =   1800
      List            =   "振替依頼書取込修正.frx":00CB
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   22
      Top             =   4500
      Width           =   2835
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
      Left            =   2100
      TabIndex        =   24
      Top             =   6600
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
      Left            =   600
      TabIndex        =   23
      Top             =   6600
      Width           =   1335
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
      ItemData        =   "振替依頼書取込修正.frx":00FB
      Left            =   1800
      List            =   "振替依頼書取込修正.frx":0108
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   480
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
      Left            =   5160
      TabIndex        =   27
      Top             =   120
      Width           =   4635
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
         TabIndex        =   28
         Top             =   420
         Width           =   3855
         Begin imText6Ctl.imText txtCIKZNO 
            DataField       =   "CIKZNO"
            DataSource      =   "dbcImportEdit"
            Height          =   285
            Left            =   1140
            TabIndex        =   17
            Top             =   1380
            Width           =   795
            _Version        =   65537
            _ExtentX        =   1402
            _ExtentY        =   503
            Caption         =   "振替依頼書取込修正.frx":0126
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "振替依頼書取込修正.frx":0192
            Key             =   "振替依頼書取込修正.frx":01B0
            MouseIcon       =   "振替依頼書取込修正.frx":01F4
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
            Height          =   615
            Left            =   1140
            TabIndex        =   44
            Top             =   960
            Width           =   2535
            Begin VB.OptionButton optCIKZSB 
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
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   180
               Width           =   675
            End
            Begin VB.OptionButton optCIKZSB 
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
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   180
               Width           =   675
            End
            Begin VB.OptionButton optCIKZSB 
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
            Begin VB.Label lblCIKZSB 
               BackColor       =   &H000000FF&
               Caption         =   "口座種別"
               DataField       =   "CIKZSB"
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
               TabIndex        =   45
               Top             =   180
               Width           =   795
            End
         End
         Begin imText6Ctl.imText txtCISITN 
            DataField       =   "CISITN"
            DataSource      =   "dbcImportEdit"
            Height          =   285
            Left            =   1200
            TabIndex        =   14
            Top             =   660
            Width           =   375
            _Version        =   65537
            _ExtentX        =   661
            _ExtentY        =   503
            Caption         =   "振替依頼書取込修正.frx":0210
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "振替依頼書取込修正.frx":027C
            Key             =   "振替依頼書取込修正.frx":029A
            MouseIcon       =   "振替依頼書取込修正.frx":02DE
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
         Begin imText6Ctl.imText txtCIBANK 
            DataField       =   "CIBANK"
            DataSource      =   "dbcImportEdit"
            Height          =   285
            Left            =   1200
            TabIndex        =   13
            Top             =   300
            Width           =   495
            _Version        =   65537
            _ExtentX        =   873
            _ExtentY        =   503
            Caption         =   "振替依頼書取込修正.frx":02FA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "振替依頼書取込修正.frx":0366
            Key             =   "振替依頼書取込修正.frx":0384
            MouseIcon       =   "振替依頼書取込修正.frx":03C8
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
            Top             =   660
            Width           =   1935
         End
      End
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
         TabIndex        =   30
         Top             =   1020
         Width           =   4035
         Begin imText6Ctl.imText txtCIYBTK 
            DataField       =   "CIYBTK"
            DataSource      =   "dbcImportEdit"
            Height          =   285
            Left            =   1860
            TabIndex        =   18
            Top             =   480
            Width           =   375
            _Version        =   65537
            _ExtentX        =   661
            _ExtentY        =   503
            Caption         =   "振替依頼書取込修正.frx":03E4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "振替依頼書取込修正.frx":0450
            Key             =   "振替依頼書取込修正.frx":046E
            MouseIcon       =   "振替依頼書取込修正.frx":04B2
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
         Begin imText6Ctl.imText txtCIYBTN 
            DataField       =   "CIYBTN"
            DataSource      =   "dbcImportEdit"
            Height          =   285
            Left            =   1860
            TabIndex        =   19
            Top             =   960
            Width           =   855
            _Version        =   65537
            _ExtentX        =   1508
            _ExtentY        =   503
            Caption         =   "振替依頼書取込修正.frx":04CE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "振替依頼書取込修正.frx":053A
            Key             =   "振替依頼書取込修正.frx":0558
            MouseIcon       =   "振替依頼書取込修正.frx":059C
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
         Begin VB.Label Label22 
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
            TabIndex        =   43
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
            TabIndex        =   42
            Top             =   480
            Width           =   1275
         End
      End
      Begin VB.OptionButton optCIKKBN 
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
      Begin VB.OptionButton optCIKKBN 
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
      Begin imText6Ctl.imText txtCIKZNM 
         DataField       =   "CIKZNM"
         DataSource      =   "dbcImportEdit"
         Height          =   285
         Left            =   420
         TabIndex        =   20
         Top             =   2580
         Width           =   3735
         _Version        =   65537
         _ExtentX        =   6588
         _ExtentY        =   503
         Caption         =   "振替依頼書取込修正.frx":05B8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "振替依頼書取込修正.frx":0624
         Key             =   "振替依頼書取込修正.frx":0642
         MouseIcon       =   "振替依頼書取込修正.frx":0686
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
         TabIndex        =   60
         Top             =   2340
         Width           =   1395
      End
      Begin VB.Label lblCIKKBN 
         BackColor       =   &H000000FF&
         Caption         =   "金融機関種別"
         DataField       =   "CIKKBN"
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
         TabIndex        =   52
         Top             =   180
         Width           =   1095
      End
   End
   Begin imNumber6Ctl.imNumber txtCISKGK 
      DataField       =   "CISKGK"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
      _Version        =   65537
      _ExtentX        =   1931
      _ExtentY        =   503
      Calculator      =   "振替依頼書取込修正.frx":06A2
      Caption         =   "振替依頼書取込修正.frx":06C2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":072E
      Keys            =   "振替依頼書取込修正.frx":074C
      MouseIcon       =   "振替依頼書取込修正.frx":0796
      Spin            =   "振替依頼書取込修正.frx":07B2
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
      Left            =   5040
      TabIndex        =   31
      Top             =   3240
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
         ItemData        =   "振替依頼書取込修正.frx":07DA
         Left            =   1500
         List            =   "振替依頼書取込修正.frx":07FF
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   32
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
         ItemData        =   "振替依頼書取込修正.frx":0841
         Left            =   3900
         List            =   "振替依頼書取込修正.frx":0866
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   34
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
         TabIndex        =   36
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
            Name            =   "ＭＳ Ｐゴシック"
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
         Bindings        =   "振替依頼書取込修正.frx":08A8
         Height          =   2040
         Left            =   120
         TabIndex        =   33
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
         Bindings        =   "振替依頼書取込修正.frx":08BE
         Height          =   2040
         Left            =   2400
         TabIndex        =   35
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
         TabIndex        =   54
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
         TabIndex        =   53
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
      Left            =   5520
      TabIndex        =   26
      Top             =   6600
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
      Left            =   4020
      TabIndex        =   25
      Top             =   6600
      Width           =   1335
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
      Left            =   8100
      TabIndex        =   0
      Top             =   6600
      Width           =   1335
   End
   Begin ORADCLibCtl.ORADC dbcImportEdit 
      Height          =   315
      Left            =   5940
      Top             =   7200
      Visible         =   0   'False
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
      _ExtentY        =   556
      _StockProps     =   207
      BackColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.01
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   "dcssvr03"
      Connect         =   "kumon/kumon"
      RecordSource    =   "select * from tchogoshaimport where 1=-1"
   End
   Begin imText6Ctl.imText txtCIKJNM 
      DataField       =   "CIKJNM"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   1920
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "振替依頼書取込修正.frx":08D6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":0942
      Key             =   "振替依頼書取込修正.frx":0960
      MouseIcon       =   "振替依頼書取込修正.frx":09A4
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
   Begin imText6Ctl.imText txtCIKYCD 
      DataField       =   "CIKYCD"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   615
      _Version        =   65537
      _ExtentX        =   1085
      _ExtentY        =   503
      Caption         =   "振替依頼書取込修正.frx":09C0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":0A2C
      Key             =   "振替依頼書取込修正.frx":0A4A
      MouseIcon       =   "振替依頼書取込修正.frx":0A8E
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
      Format          =   ""
      FormatMode      =   1
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
   Begin imText6Ctl.imText txtCIHGCD 
      DataField       =   "CIHGCD"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   1560
      Width           =   495
      _Version        =   65537
      _ExtentX        =   873
      _ExtentY        =   503
      Caption         =   "振替依頼書取込修正.frx":0AAA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":0B16
      Key             =   "振替依頼書取込修正.frx":0B34
      MouseIcon       =   "振替依頼書取込修正.frx":0B78
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
   Begin ORADCLibCtl.ORADC dbcItakushaMaster 
      Height          =   315
      Left            =   4080
      Top             =   7200
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
   Begin imText6Ctl.imText txtCISTNM 
      DataField       =   "CISTNM"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   2640
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "振替依頼書取込修正.frx":0B94
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":0C00
      Key             =   "振替依頼書取込修正.frx":0C1E
      MouseIcon       =   "振替依頼書取込修正.frx":0C62
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
   Begin imText6Ctl.imText txtCIKNNM 
      DataField       =   "CIKNNM"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   2280
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "振替依頼書取込修正.frx":0C7E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":0CEA
      Key             =   "振替依頼書取込修正.frx":0D08
      MouseIcon       =   "振替依頼書取込修正.frx":0D4C
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
   Begin imText6Ctl.imText txtCIKSCD 
      DataField       =   "CIKSCD"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   375
      _Version        =   65537
      _ExtentX        =   661
      _ExtentY        =   503
      Caption         =   "振替依頼書取込修正.frx":0D68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":0DD4
      Key             =   "振替依頼書取込修正.frx":0DF2
      MouseIcon       =   "振替依頼書取込修正.frx":0E36
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
   Begin imText6Ctl.imText txtCIBKNM 
      DataField       =   "CIBKNM"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   3360
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "振替依頼書取込修正.frx":0E52
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":0EBE
      Key             =   "振替依頼書取込修正.frx":0EDC
      MouseIcon       =   "振替依頼書取込修正.frx":0F20
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
   Begin imText6Ctl.imText txtCISINM 
      DataField       =   "CISINM"
      DataSource      =   "dbcImportEdit"
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   3720
      Width           =   2835
      _Version        =   65537
      _ExtentX        =   5001
      _ExtentY        =   503
      Caption         =   "振替依頼書取込修正.frx":0F3C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "振替依頼書取込修正.frx":0FA8
      Key             =   "振替依頼書取込修正.frx":0FC6
      MouseIcon       =   "振替依頼書取込修正.frx":100A
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
   Begin VB.Label lblCIMUPD 
      BackColor       =   &H000000FF&
      Caption         =   "反映なし"
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
      Left            =   3120
      TabIndex        =   79
      Top             =   7500
      Width           =   855
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
      Left            =   2160
      TabIndex        =   70
      Top             =   7500
      Width           =   855
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
      Left            =   480
      TabIndex        =   77
      Top             =   7500
      Width           =   795
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
      TabIndex        =   76
      Top             =   4185
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
      TabIndex        =   75
      Top             =   4140
      Width           =   555
      WordWrap        =   -1  'True
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
      Left            =   1320
      TabIndex        =   72
      Top             =   7500
      Width           =   795
   End
   Begin VB.Label Label10 
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
      TabIndex        =   71
      Top             =   4560
      Width           =   1395
   End
   Begin VB.Image imgCIWMSG 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Left            =   240
      Picture         =   "振替依頼書取込修正.frx":1026
      Top             =   4860
      Visible         =   0   'False
      Width           =   480
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
      Left            =   600
      TabIndex        =   69
      Top             =   7200
      Visible         =   0   'False
      Width           =   3300
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
      Left            =   285
      TabIndex        =   68
      Top             =   3390
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
      Left            =   285
      TabIndex        =   67
      Top             =   3765
      Width           =   1395
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
      Left            =   1800
      TabIndex        =   66
      Top             =   60
      Width           =   1755
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCIITKB 
      BackColor       =   &H000000FF&
      Caption         =   "委託者区分"
      DataField       =   "CIITKB"
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
      Left            =   3720
      TabIndex        =   65
      Top             =   540
      Width           =   975
   End
   Begin VB.Label Label7 
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
      Left            =   3615
      TabIndex        =   64
      Top             =   120
      Width           =   180
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
      Left            =   3840
      TabIndex        =   63
      Top             =   60
      Width           =   795
      WordWrap        =   -1  'True
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
      Left            =   285
      TabIndex        =   62
      Top             =   105
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
      Height          =   210
      Left            =   285
      TabIndex        =   61
      Top             =   2325
      Width           =   1395
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
      Height          =   210
      Left            =   285
      TabIndex        =   59
      Top             =   2685
      Width           =   1395
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
      Left            =   7800
      TabIndex        =   58
      Top             =   7200
      Width           =   795
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
      Left            =   8820
      TabIndex        =   57
      Top             =   7200
      Width           =   915
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
      Height          =   210
      Left            =   285
      TabIndex        =   56
      Top             =   1275
      Width           =   1395
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
      Height          =   210
      Left            =   285
      TabIndex        =   55
      Top             =   525
      Width           =   1395
   End
   Begin VB.Label Label17 
      Alignment       =   1  '右揃え
      Caption         =   "振替金額"
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
      Left            =   285
      TabIndex        =   41
      Top             =   3045
      Width           =   1395
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
      TabIndex        =   40
      Top             =   900
      Width           =   2355
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
      Height          =   210
      Left            =   285
      TabIndex        =   39
      Top             =   885
      Width           =   1395
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
      Height          =   210
      Left            =   285
      TabIndex        =   38
      Top             =   1605
      Width           =   1395
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
      Height          =   210
      Left            =   285
      TabIndex        =   37
      Top             =   1965
      Width           =   1395
   End
   Begin VB.Menu mnuFile 
      Caption         =   "ﾌｧｲﾙ(&F)"
      Begin VB.Menu mnuEnd 
         Caption         =   "戻る(&B)"
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
Private mErrMsgOn As Boolean
Private mCaption As String
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
    Call mForm.LockedControl(False) '//常にデータは編集可能にしておく
'    cmdUpdate.Enabled = blMode
'    cmdCancel.Enabled = blMode
    'mForm.LockedControl() で警告表示が赤色の為、消える！
    lblERRMSG.Visible = True
    '//2007/06/07 口座名義人は常に入力しない：保護者名(カナ)をコピーする様に仕様変更
    txtCIKZNM.Enabled = False
    lblKouzaName.Enabled = False
    cmdKakutei.Enabled = Not blMode
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
        lblCIITKB.Caption = cboABKJNM.ItemData(cboABKJNM.ListIndex)
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
    Case mRimp.errWarning
        If cboCIOKFG.ItemData(cboCIOKFG.ListIndex) = mRimp.updNormal Then
            '//再チェック時に警告に戻るので選択の意味が無い
            Call MsgBox("「警告」データを反映するには" & vbCrLf & "「" & mRimp.mUpdateMessage(mRimp.updWarnUpd) & "」を選択してください。", vbInformation + vbOKOnly, mCaption)
            '//元に戻す
            cboCIOKFG.ListIndex = Val(lblCIOKFG.Caption) + 2    '// -2 ～ 2
            Exit Sub
        Else
            If InStr(lblCIWMSG.Caption, "解約状態") Then
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
    '//キーを押した時に更新可能か判断
'    cmdUpdate.Enabled = mCheckUpdate    '//更新ボタンの制御：データ表示時にイベントが発生しても可能なように！
'    cmdCancel.Enabled = cmdUpdate.Enabled
    Call SendKeys("{TAB}")  '//結果を正しく見せたいのでフォーカス移動
End Sub

Private Sub cboShitenYomi_Click()
    If dblBankList.Text = "" Then
        Exit Sub
    End If
    Call gdDBS.BankDbListRefresh(dbcShiten, cboShitenYomi, dblShitenList, eBankRecordKubun.Shiten, Left(dblBankList.Text, 4))
    cmdKakutei.Enabled = False
End Sub

Private Sub chkCIMUPD_Click()
    lblCIMUPD.Caption = Val(chkCIMUPD.Value)
    Call pLockedControl(True)
    Call pButtonControl(True)
End Sub

Private Sub cmdCancel_Click()
    Call dbcImportEdit.UpdateControls
    'Call cboABKJNM.SetFocus
    Call pLockedControl(False)
    Call lblCIERROR_Change
    Call pButtonControl(False)
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

Private Sub cmdKakutei_Click()
    If dblBankList.Text = "" Or dblShitenList.Text = "" Then
        Exit Sub
    End If
    txtCiBANK.Text = Left(dblBankList.Text, 4)
    txtCISITN.Text = Left(dblShitenList.Text, 3)
    '//入力された金融機関名＆支店名を強制書き換え
    txtCIBKNM.Text = Mid(dblBankList.Text, 6)
    lblBankName.Caption = Mid(dblBankList.Text, 6)
    txtCISINM.Text = Mid(dblShitenList.Text, 5)
    lblShitenName.Caption = Mid(dblShitenList.Text, 5)
    cmdKakutei.Enabled = False
'//2006/08/22 確定後交信可能に！
    Call pLockedControl(True)
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
    If txtCIKJNM.Text = "" Then
        Set obj = txtCIKJNM
        msg = "保護者名(漢字)は必須入力です."
    ElseIf False = str.CheckLength(txtCIKJNM.Text) Then
        Set obj = txtCIKJNM
        msg = "保護者名(漢字)に半角が含まれています."
    End If
    '//保護者・カナ名称は必須
    '//2007/06/07 必須 復活：口座名義人と同じ値とする為
    If txtCIKNNM.Text = "" Then
        Set obj = txtCIKNNM
        msg = "保護者名(カナ)は必須入力です."
    ElseIf False = str.CheckLength(txtCIKNNM.Text, vbNarrow) Then
        Set obj = txtCIKNNM
        msg = "保護者名(カナ)に全角が含まれています."
    ElseIf 0 < InStr(txtCIKNNM.Text, "ｰ") Then
        Set obj = txtCIKNNM
        msg = "保護者名(カナ)に長音が含まれています."
    End If
#If 0 Then  '//項目なし
    If IsNull(txtCIKYxx(1).Number) Then
        Set obj = txtCIKYxx(1)
        msg = "契約期間の終了日は必須入力です."
    ElseIf txtCIKYxx(0).Text > txtCIKYxx(1).Text Then
        Set obj = txtCIKYxx(0)
        msg = "契約期間が不正です."
    ElseIf IsNull(txtCIFKxx(1).Number) Then
        Set obj = txtCIFKxx(1)
        msg = "振替期間の終了日は必須入力です."
    ElseIf txtCIFKxx(0).Text > txtCIFKxx(1).Text Then
        Set obj = txtCIFKxx(0)
        msg = "振替期間が不正です."
    End If
#End If
    If lblCIKKBN.Caption = eBankKubun.KinnyuuKikan Then
        If txtCiBANK.Text = "" Or lblBankName.Caption = "" Then
            Set obj = txtCiBANK
            msg = "金融機関は必須入力です."
        ElseIf txtCISITN.Text = "" Or lblShitenName.Caption = "" Then
            Set obj = txtCISITN
            msg = "支店は必須入力です."
        ElseIf Not (lblCIKZSB.Caption = eBankYokinShubetsu.Futsuu _
                 Or lblCIKZSB.Caption = eBankYokinShubetsu.Touza) Then
            Set obj = optCIKZSB(eBankYokinShubetsu.Futsuu)
            msg = "預金種別は必須入力です."
        ElseIf txtCIKZNO.Text = "" Then
            Set obj = txtCIKZNO
            msg = "口座番号は必須入力です."
        End If
    ElseIf lblCIKKBN.Caption = eBankKubun.YuubinKyoku Then
        If txtCIYBTK.Text = "" Then
            Set obj = txtCIYBTK
            msg = "通帳記号は必須入力です."
        ElseIf txtCIYBTN.Text = "" Then
            Set obj = txtCIYBTN
            msg = "通帳番号は必須入力です."
        ElseIf "1" <> Right(txtCIYBTN.Text, 1) Then
'//2006/04/26 末尾番号チェック
            Set obj = txtCIYBTN
            msg = "通帳番号の末尾が「１」以外です."
        End If
    End If
    If txtCIKZNM.Text = "" Then
        Set obj = txtCIKZNM
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
    mErrMsgOn = False
    Call txtCIKYCD_KeyDown(vbKeyReturn, 0)
    mErrMsgOn = True
'    cmdUpdate.Enabled = False
'    cmdCancel.Enabled = False
    'Call dbcImportEdit.UpdateControls
    Call pButtonControl(False, True)
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
    cboABKJNM.BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCIITKB.DataField & "E"))
    '//金融機関区分のエラー表示
    optCIKKBN(0).BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCIKKBN.DataField & "E"), False)
    optCIKKBN(1).BackColor = optCIKKBN(0).BackColor
    '//預金種別のエラー表示
    optCIKZSB(0).BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCIKZSB.DataField & "E"), False)
    optCIKZSB(1).BackColor = optCIKZSB(0).BackColor
    optCIKZSB(2).BackColor = optCIKZSB(0).BackColor
End Function

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
    cboABKJNM.BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCIITKB.DataField & "E"))
    '//金融機関区分のエラー表示
    optCIKKBN(0).BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCIKKBN.DataField & "E"), False)
    optCIKKBN(1).BackColor = optCIKKBN(0).BackColor
    '//預金種別のエラー表示
    optCIKZSB(0).BackColor = mRimp.ErrorStatus(dbcImportEdit.Recordset.Fields(lblCIKZSB.DataField & "E"), False)
    optCIKZSB(1).BackColor = optCIKZSB(0).BackColor
    optCIKZSB(2).BackColor = optCIKZSB(0).BackColor
    cboCIOKFG.ListIndex = Val(lblCIOKFG.Caption) + 2    '// -2 ～ 2
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

Private Sub lblCIERROR_Change()
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
End Sub

Private Sub lblCIITKB_Change()
    Select Case lblCIITKB.Caption
    Case 0:     cboABKJNM.ListIndex = lblCIITKB.Caption
    Case 1:     cboABKJNM.ListIndex = lblCIITKB.Caption
    Case Else:  cboABKJNM.ListIndex = -1
    End Select
End Sub

Private Sub lblCIKKBN_Change()
'    On Error Resume Next
    '//ブランクはエラーとする
    If Not IsNull(lblCIKKBN.Caption) And "" <> lblCIKKBN.Caption Then
        optCIKKBN(lblCIKKBN.Caption).Value = True
    End If
End Sub

Private Sub lblCIKZSB_Change()
    If Not IsNull(lblCIKZSB.Caption) And "" <> lblCIKZSB.Caption Then
        optCIKZSB(Val(lblCIKZSB.Caption)).Value = True
    Else
'//設定すると更新フラグが立ってしまうので止める
'//        optCIKZSB(0).Value = True
    End If
End Sub

Private Sub lblCIMUPD_Change()
    chkCIMUPD.Value = Abs(Val(lblCIMUPD.Caption) <> 0)
End Sub

Private Sub lblCIOKFG_Change()
'''    If Not IsNull(lblCIOKFG.Caption) And "" <> lblCIOKFG.Caption Then
'''        cboCIOKFG.ListIndex = Val(lblCIOKFG.Caption) + 2    '// -2 ～ 2
'''    End If
End Sub

Private Sub lblCIWMSG_Change()
    txtCIWMSG.Text = lblCIWMSG.Caption
    imgCIWMSG.Visible = lblCIWMSG.Caption <> ""
End Sub

Private Sub mnuEnd_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

Private Sub optCIKKBN_Click(Index As Integer)
    fraKinnyuuKikan.Tag = Index
    Call fraBank(Index).ZOrder(0)
    fraBankList.Visible = Index = 0
    lblCIKKBN.Caption = Index
    '//フォーカスが消えるので設定する.
    txtCiBANK.TabStop = Index = eBankKubun.KinnyuuKikan
    txtCISITN.TabStop = Index = eBankKubun.KinnyuuKikan
    txtCIKZNO.TabStop = Index = eBankKubun.KinnyuuKikan
    txtCIYBTK.TabStop = Index = eBankKubun.YuubinKyoku
    txtCIYBTN.TabStop = Index = eBankKubun.YuubinKyoku
'    cmdUpdate.Enabled = True
'    cmdCancel.Enabled = True
    Call pButtonControl(True)
End Sub

Private Sub optCIKZSB_Click(Index As Integer)
    lblCIKZSB.Caption = Index
'    cmdUpdate.Enabled = True
'    cmdCancel.Enabled = True
    Call pButtonControl(True)
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

Private Sub txtCIHGCD_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCIKJNM_Change()
    If Len(Trim(txtCIKJNM.Text)) = 0 Then
        txtCIKNNM.Text = ""
        txtCIKZNM.Text = ""
    End If
    Call pButtonControl(True)
End Sub

Private Sub txtCIKJNM_Furigana(Yomi As String)
'//2007/06/07 カナ名と口座名義人名が同じ
'    '//現在の読みカナ名と口座名義人名が同じなら読みカナ名と口座名義人名に転送
'    If Trim(txtCIKNNM.Text) = Trim(txtCIKZNM.Text) Then
'        txtCIKNNM.Text = txtCIKNNM.Text & Yomi
'        txtCIKZNM.Text = txtCIKNNM.Text
'    Else
'        txtCIKNNM.Text = txtCIKNNM.Text & Yomi
'    End If
     txtCIKNNM.Text = txtCIKNNM.Text & Yomi
     txtCIKZNM.Text = txtCIKNNM.Text
End Sub

Private Sub txtCIKNNM_Change()
    txtCIKZNM.Text = txtCIKNNM.Text '//2007/06/07 保護者名(カナ)＝口座名義人名
    Call pButtonControl(True)
End Sub

Private Sub txtCIKSCD_Change()
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
        
    If "" = Trim(txtCiKYCD.Text) Then
        Exit Sub
    End If
'//2002/12/10 教室区分(??KSCD)は使用しない
'//    sql = "SELECT DISTINCT BAITKB,BAKYCD,BAKSCD,BAKJNM FROM tbKeiyakushaMaster"
    sql = "SELECT DISTINCT BAITKB,BAKYCD,BAKJNM,BAKYED FROM tbKeiyakushaMaster"
    sql = sql & " WHERE BAITKB = '" & lblCIITKB.Caption & "'"
    sql = sql & "   AND BAKYCD = '" & txtCiKYCD.Text & "'"
'//2006/03/31 解約状態を表示するように変更
'    sql = sql & "   AND TO_CHAR(SYSDATE,'YYYYMMDD') BETWEEN BAKYST AND BAKYED" '//有効データ絞込み
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If 0 = dyn.RecordCount Then
        Call dyn.Close
        KeyCode = 0
        If mErrMsgOn = True Then
            '//                                        「契約者番号」
            Call MsgBox("該当データは存在しません.(" & lblKeiyakushaCode.Caption & ")", vbInformation, mCaption)
            Call txtCiKYCD.SetFocus
        End If
        Exit Sub
    End If
#If 1 Then
    lblBAKJNM.Caption = dyn.Fields("BAKJNM")
#Else
    lblBAKJNM.Caption = IIf(dyn.Fields("BAKYED") < gdDBS.sysDate("yyyymmdd"), "(解約)", "") & _
                        dyn.Fields("BAKJNM")
#End If
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
    Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Shiten, txtCiBANK.Text, txtCISITN.Text, vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"支店名_漢字" で読めない
    Set dyn = Nothing
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

Private Sub txtCISITN_LostFocus()
    If 0 <= Len(Trim(txtCISITN.Text)) And Len(Trim(txtCISITN.Text)) < 3 Then
        lblShitenName.Caption = ""
        Exit Sub
    End If
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Set dyn = gdDBS.SelectBankMaster("DAKJNM", eBankRecordKubun.Shiten, Trim(txtCiBANK.Text), Trim(txtCISITN.Text), vDate:=gdDBS.sysDate("YYYYMMDD"))
    lblShitenName.Caption = gdDBS.Nz(dyn.Fields("DAKJNM"))   '//"支店名_漢字" で読めない
End Sub

Private Sub txtCISKGK_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCISTNM_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCIYBTK_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCIYBTN_Change()
    Call pButtonControl(True)
End Sub
