VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMainMenu 
   Caption         =   "メインメニュー"
   ClientHeight    =   4530
   ClientLeft      =   2145
   ClientTop       =   2355
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   302
   ScaleMode       =   3  'ﾋﾟｸｾﾙ
   ScaleWidth      =   466
   Begin VB.Timer tmrTimer 
      Interval        =   1000
      Left            =   4800
      Top             =   3960
   End
   Begin VB.Frame fraSysDate 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   375
      Left            =   5580
      TabIndex        =   14
      Top             =   0
      Width           =   1155
      Begin VB.Label lblSysDate 
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   60
         Width           =   855
      End
   End
   Begin TabDlg.SSTab tabMenu 
      Height          =   3795
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   6694
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "月例処理"
      TabPicture(0)   =   "メインメニュー.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdMenu(11)"
      Tab(0).Control(1)=   "cmdMenu(5)"
      Tab(0).Control(2)=   "cmdMenu(10)"
      Tab(0).Control(3)=   "cmdMenu(4)"
      Tab(0).Control(4)=   "cmdMenu(3)"
      Tab(0).Control(5)=   "cmdMenu(7)"
      Tab(0).Control(6)=   "cmdMenu(8)"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "取込処理"
      TabPicture(1)   =   "メインメニュー.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdMenu(9)"
      Tab(1).Control(1)=   "cmdMenu(6)"
      Tab(1).Control(2)=   "cmdMenu(2)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "メンテナンス"
      TabPicture(2)   =   "メインメニュー.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "cmdMenu(103)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdMenu(102)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdMenu(104)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdMenu(105)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdMenu(101)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdMenu(107)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdMenu(108)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "システム情報"
      TabPicture(3)   =   "メインメニュー.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdMenu(106)"
      Tab(3).ControlCount=   1
      Begin VB.CommandButton cmdMenu 
         Caption         =   "保護者マスタ履歴 照会"
         Height          =   495
         Index           =   108
         Left            =   3480
         TabIndex        =   27
         Top             =   1800
         Width           =   2355
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "口座振替依頼書(履歴)"
         Height          =   495
         Index           =   11
         Left            =   -71520
         TabIndex        =   20
         Top             =   540
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "オーナーマスタチェックリスト"
         Height          =   495
         Index           =   107
         Left            =   3480
         TabIndex        =   19
         Top             =   1200
         Width           =   2355
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "オーナーマスタデータ作成"
         Height          =   495
         Index           =   5
         Left            =   -71520
         TabIndex        =   18
         Top             =   1740
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "口座振替依頼書(入力)"
         Height          =   495
         Index           =   10
         Left            =   -74340
         TabIndex        =   17
         Top             =   540
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "振込依頼書(取込)"
         Height          =   495
         Index           =   9
         Left            =   -74340
         TabIndex        =   16
         Top             =   540
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "金融機関データ取込"
         Height          =   495
         Index           =   6
         Left            =   -74340
         TabIndex        =   13
         Top             =   2940
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "振替予定表 兼 解約通知書(取込)"
         Height          =   495
         Index           =   2
         Left            =   -74340
         TabIndex        =   12
         Top             =   1140
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "振替予定表(累積)"
         Height          =   495
         Index           =   4
         Left            =   -71520
         TabIndex        =   11
         Top             =   2940
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "口座振替データ作成"
         Height          =   495
         Index           =   3
         Left            =   -74340
         TabIndex        =   10
         Top             =   1740
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "口座振替依頼書(印刷)"
         Height          =   495
         Index           =   7
         Left            =   -74340
         TabIndex        =   9
         Top             =   1140
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "振替予定表(印刷)"
         Height          =   495
         Index           =   8
         Left            =   -74340
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   8
         Top             =   2340
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "委託者マスタメンテナンス"
         Height          =   495
         Index           =   101
         Left            =   660
         TabIndex        =   7
         Top             =   540
         Width           =   2355
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "休日マスタメンテナンス"
         Height          =   495
         Index           =   105
         Left            =   660
         TabIndex        =   6
         Top             =   3000
         Width           =   2355
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "金融機関マスタメンテナンス"
         Height          =   495
         Index           =   104
         Left            =   660
         TabIndex        =   5
         Top             =   2400
         Width           =   2355
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "オーナーマスタメンテナンス"
         Height          =   495
         Index           =   102
         Left            =   660
         TabIndex        =   4
         Top             =   1200
         Width           =   2355
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "保護者マスタメンテナンス"
         Height          =   495
         Index           =   103
         Left            =   660
         TabIndex        =   3
         Top             =   1800
         Width           =   2355
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "基本情報登録"
         Height          =   495
         Index           =   106
         Left            =   -74340
         TabIndex        =   2
         Top             =   540
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "終了(&X)"
      Height          =   495
      Index           =   0
      Left            =   5340
      TabIndex        =   0
      Top             =   3900
      Width           =   1335
   End
   Begin VB.Frame fraTimer 
      BorderStyle     =   0  'なし
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   3495
      Begin VB.Label lblClientTime 
         Caption         =   "2007/06/13 13:58:11"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1365
         TabIndex        =   25
         Top             =   180
         Width           =   1995
      End
      Begin VB.Label lblServerTime 
         Caption         =   "2007/06/13 13:58:11"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1365
         TabIndex        =   24
         Top             =   390
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   1  '右揃え
         AutoSize        =   -1  'True
         Caption         =   "Client Time："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   105
         TabIndex        =   23
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   1  '右揃え
         AutoSize        =   -1  'True
         Caption         =   "Server Time："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   60
         TabIndex        =   22
         Top             =   390
         Width           =   1200
      End
   End
   Begin VB.Label lblLoginUserName 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      TabIndex        =   26
      Top             =   4320
      Width           =   3555
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
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mForm As New FormClass
Private mReg As New RegistryClass

Private Enum eButton
    eEnd = 0
    '//Left Menu
    eFrmFurikaeIraishoInput = 10        '//振替依頼書(入力)
'//2012/07/05 振替依頼書(履歴)追加
    eFrmFurikaeIraishoRireki = 11       '//振替依頼書(履歴)
    eFrmKouzaFurikaeIraishoPrint = 7    '//振替依頼書プリント
    eFrmFurikaeYoteiPrint = 8           '//振替予定プリント
    eFrmKouzaFurikaeExportYotei = 1     '//口座振替テキスト出力（予定）
    eFrmFurikaeYoteiImport = 2          '//振替予定取込み
    eFrmKouzaFurikaeExportJisseki = 3   '//口座振替テキスト出力（請求）
    eFrmFurikaeDataRuiseki = 4
'//2012/07/05 振込依頼書取込み（保護者マスタ）追加
    eFrmFurikaeReqImport = 9            '//振込依頼書取込み（保護者マスタ）
    '//Right Menu-2
    eFrmKeiyakushaMasterExport = 5      '//契約者テキスト出力
    eFrmBankDataImport = 6              '//金融機関取込み
    '//Right Menu
    eFrmItakushaMaster = 101            '//委託者マスタメンテ
    eFrmKeiyakushaMaster = 102          '//契約者マスタメンテ
    eFrmHogoshaMaster = 103             '//保護者マスタメンテ
    eFrmBankMaster = 104                '//金融機関マスタメンテ
    eFrmHolidayMaster = 105             '//休日マスタメンテ
    eFrmSystemInfomation = 106          '//基本情報
    eFrmKeiyakushaCheckList = 107       '//オーナーマスタチェックリスト：2007/02/05 追加 ＷＡＯ
    efrmHogoshaMasterRireki = 108       '//保護者履歴：照会 2012/07/09
End Enum

Private Sub cmdMenu_Click(Index As Integer)
    Dim frm As Form
    Select Case Index
    Case eButton.eEnd
        Unload Me       'Unload()にデストラクタあり
    Case eButton.eFrmItakushaMaster
        Set frm = frmItakushaMaster
    Case eButton.eFrmHogoshaMaster, eButton.eFrmFurikaeIraishoInput         '//eFrmHogoshaMaster
        Set frm = frmHogoshaMaster
    Case eButton.eFrmFurikaeIraishoRireki   '//振替依頼書(履歴)
        Set frm = frmHogoshaMasterRireki
    Case eButton.eFrmFurikaeYoteiImport
        Set frm = frmFurikaeYoteiImport
    Case eButton.eFrmFurikaeDataRuiseki
        Set frm = frmFurikaeDataRuiseki
    Case eButton.eFrmKeiyakushaMaster
        Set frm = frmKeiyakushaMaster
    Case eFrmKeiyakushaCheckList
        Set frm = frmKeiyakushaCheckList
    Case eButton.eFrmSystemInfomation
        Set frm = frmSystemInfomation
    Case eButton.eFrmHolidayMaster
        Set frm = frmHolidayMaster
    Case eButton.eFrmKouzaFurikaeExportYotei
        Set frm = frmKouzaFurikaeExport
        frm.chkJisseki.Value = 0
    Case eButton.eFrmKouzaFurikaeExportJisseki
        Set frm = frmKouzaFurikaeExport
        frm.chkJisseki.Value = 1
    Case eButton.eFrmBankMaster
        Set frm = frmBankMaster
    Case eButton.eFrmKeiyakushaMasterExport
        Set frm = frmKeiyakushaMasterExport
    Case eButton.eFrmBankDataImport
        Set frm = frmBankDataImport
    Case eButton.eFrmKouzaFurikaeIraishoPrint
        Set frm = frmKouzaFurikaeIraishoPrint
    Case eButton.eFrmFurikaeYoteiPrint
        Set frm = frmFurikaeYoteiPrint
    Case eButton.eFrmFurikaeReqImport
        Set frm = frmFurikaeReqImport
    Case eButton.efrmHogoshaMasterRireki       '//保護者履歴：照会 2012/07/09
        Set frm = frmHogoshaMasterRireki
    End Select
    '//ボタンを押した時のみ記憶する
    mReg.MenuButton = Index
    mReg.MenuTab = tabMenu.Tab
    If UCase(TypeName(frm)) <> UCase("Nothing") Then
        Set gdForm = Me
        Call frm.Show
        Call Me.Hide
    End If
End Sub

Private Sub Form_Activate()
    '//SetFocus 出来ない時のエラー対応
    On Error Resume Next
    Call cmdMenu(mReg.MenuButton).SetFocus
End Sub

Private Sub Form_Load()
    Call mForm.Init(Me, gdDBS)
'    cmdMenu(eButton.eFrmFurikaeYoteiImport).Caption = " 振替金額予定表" & vbCrLf & "兼 解約通知書 (取込)"
'    cmdMenu(eButton.eFrmFurikaeDataRuiseki).Caption = " 振替金額予定表" & vbCrLf & "兼 解約通知書 (累積)"
    Call mForm.MoveSysDate
    tabMenu.Tab = mReg.MenuTab
    
    tmrTimer.Interval = 60000    '// １秒＝1,000 / １分＝60,000
    Call tmrTimer_Timer
    Dim min As Integer
    min = DateDiff("n", CVDate(lblClientTime.Caption), CVDate(lblServerTime.Caption))
    tmrTimer.Enabled = mReg.CheckTimer() <= Abs(min)
    fraTimer.Visible = tmrTimer.Enabled
'    If tmrTimer.Enabled = True Then
'    End If
    lblLoginUserName.Caption = gdDBS.LoginUserName()
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
    Call mForm.MoveSysDate
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMainMenu = Nothing
    Set mForm = Nothing
    Call gkAllEnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub mnuEnd_Click()
    Call cmdMenu_Click(eButton.eEnd)
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

Private Sub tabMenu_Click(PreviousTab As Integer)
    '//SetFocus 出来ない時のエラー対応
    On Error Resume Next
    Call cmdMenu(mReg.MenuButton).SetFocus
End Sub

Private Sub tmrTimer_Timer()
    lblClientTime.Caption = Format(Now(), "yyyy/MM/dd HH:nn:ss")
    lblServerTime.Caption = gdDBS.sysDate("yyyy/mm/dd hh24:mi:ss")
End Sub
