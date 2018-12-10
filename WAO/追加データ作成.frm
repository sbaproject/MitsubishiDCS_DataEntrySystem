VERSION 5.00
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Begin VB.Form frmMakeNewData 
   Caption         =   "履歴データ追加"
   ClientHeight    =   3300
   ClientLeft      =   3750
   ClientTop       =   1800
   ClientWidth     =   6345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6345
   Begin VB.CommandButton cmdReturn 
      Caption         =   "上書き(&U)"
      Height          =   435
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      ToolTipText     =   "履歴を追加しないでそのまま更新する場合"
      Top             =   2580
      Width           =   1395
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "追加(&A)"
      Height          =   435
      Index           =   1
      Left            =   480
      TabIndex        =   1
      ToolTipText     =   "履歴を追加する場合"
      Top             =   2580
      Width           =   1395
   End
   Begin VB.CommandButton cmdReturn 
      Cancel          =   -1  'True
      Caption         =   "中止(&C)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   4740
      TabIndex        =   0
      ToolTipText     =   "この作業を中止して再度もとの画面を編集する場合"
      Top             =   2580
      Width           =   1335
   End
   Begin imDate6Ctl.imDate txtKeiyakuEnd 
      DataField       =   "BAKYED"
      Height          =   315
      Left            =   3540
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   556
      Calendar        =   "追加データ作成.frx":0000
      Caption         =   "追加データ作成.frx":0186
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "追加データ作成.frx":01F4
      Keys            =   "追加データ作成.frx":0212
      MouseIcon       =   "追加データ作成.frx":0270
      Spin            =   "追加データ作成.frx":028C
      AlignHorizontal =   2
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   255
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
   Begin imDate6Ctl.imDate txtFurikaeEnd 
      DataField       =   "BAFKED"
      Height          =   315
      Left            =   3780
      TabIndex        =   5
      Top             =   1920
      Width           =   855
      _Version        =   65537
      _ExtentX        =   1508
      _ExtentY        =   556
      Calendar        =   "追加データ作成.frx":02B4
      Caption         =   "追加データ作成.frx":043A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "追加データ作成.frx":04A8
      Keys            =   "追加データ作成.frx":04C6
      MouseIcon       =   "追加データ作成.frx":0524
      Spin            =   "追加データ作成.frx":0540
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
      Text            =   "2012/12"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   41253
      CenturyMode     =   0
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label19"
      Height          =   255
      Left            =   4860
      TabIndex        =   10
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H000000FF&
      Caption         =   "追加されるデータの"
      Height          =   255
      Left            =   900
      TabIndex        =   9
      Top             =   3060
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "追加されるデータの"
      Height          =   255
      Left            =   1140
      TabIndex        =   8
      Top             =   1980
      Width           =   1515
   End
   Begin VB.Label lblFurikomi 
      Alignment       =   1  '右揃え
      Caption         =   "振込開始月"
      Height          =   255
      Left            =   2700
      TabIndex        =   7
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label Label19 
      Alignment       =   1  '右揃え
      BackColor       =   &H000000FF&
      Caption         =   "有効開始日"
      Height          =   255
      Left            =   2460
      TabIndex        =   6
      Top             =   3060
      Width           =   975
   End
   Begin VB.Label lblMessage 
      Caption         =   "lblMessage"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   180
      TabIndex        =   3
      Top             =   300
      Width           =   6015
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
Attribute VB_Name = "frmMakeNewData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm As New FormClass
Private mCaption As String

'//戻りフォームで参照する変数
Public mPushButton As Integer
Public Enum ePushButton
    Cancel = 0
    Add = 1
    Update = 2
End Enum
Public mKeiyakuEnd As Long
Public mFurikaeEnd As Long

Private Sub cmdReturn_Click(Index As Integer)
'    lblPushButton.Caption = Index   '//オブジェクトを作成しても閉じるときに破棄される
    mPushButton = Index             '//こうしたら変数は相手-Form に変更した状態で見える
 '''//2002/10/18 そのままの日付とする
'''   '//年月のみの入力なので 2/31 とかが存在するため
'''    mKeiyakuEnd = Format(DateSerial(txtKeiyakuEnd.Year, txtKeiyakuEnd.Month, 1), "yyyymmdd")
'''    mFurikaeEnd = Format(DateSerial(txtFurikaeEnd.Year, txtFurikaeEnd.Month, 1), "yyyymmdd")

'/////////////////////////////////////////////////////
'//2012/12/10 注意！！！
'// オーナーマスタには契約期間のみ が存在する
'// 保護者  マスタには振替期間のみ が存在する
    mKeiyakuEnd = Left(Trim(CStr(txtKeiyakuEnd.Number)), 6) & "01"
    mFurikaeEnd = Left(Trim(CStr(txtFurikaeEnd.Number)), 6) & "01"
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    Call mForm.LockedControl(False)
'    Call Me.Move((Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2)
    Me.Height = 4000    '//スタートメニューに左右されてサイズがおかしくなるので強制的に設定する.
    Me.Icon = frmAbout.Icon
    txtKeiyakuEnd.Number = gdDBS.sysDate("YYYYMMDD")
    txtFurikaeEnd.Number = gdDBS.sysDate("YYYYMMDD")
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMakeNewData = Nothing
    Set mForm = Nothing
'    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub mnuEnd_Click()
    Call cmdReturn_Click(ePushButton.Cancel)
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

Private Sub txtFurikaeEnd_DropOpen(NoDefault As Boolean)
    txtFurikaeEnd.Calendar.Holidays = gdDBS.Holiday(txtFurikaeEnd.Year)
End Sub

Private Sub txtKeiyakuEnd_DropOpen(NoDefault As Boolean)
    txtKeiyakuEnd.Calendar.Holidays = gdDBS.Holiday(txtKeiyakuEnd.Year)
End Sub

