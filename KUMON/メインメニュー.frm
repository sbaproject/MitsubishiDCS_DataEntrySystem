VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMainMenu 
   Caption         =   "���C�����j���["
   ClientHeight    =   4635
   ClientLeft      =   2145
   ClientTop       =   2355
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   309
   ScaleMode       =   3  '�߸��
   ScaleWidth      =   466
   Begin VB.Timer tmrTimer 
      Interval        =   1000
      Left            =   4800
      Top             =   3960
   End
   Begin VB.Frame fraSysDate 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      Height          =   375
      Left            =   5580
      TabIndex        =   16
      Top             =   0
      Width           =   1155
      Begin VB.Label lblSysDate 
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   17
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
      TabCaption(0)   =   "���Ꮘ��"
      TabPicture(0)   =   "���C�����j���[.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdMenu(10)"
      Tab(0).Control(1)=   "cmdMenu(3)"
      Tab(0).Control(2)=   "cmdMenu(4)"
      Tab(0).Control(3)=   "cmdMenu(1)"
      Tab(0).Control(4)=   "cmdMenu(7)"
      Tab(0).Control(5)=   "cmdMenu(8)"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "�捞����"
      TabPicture(1)   =   "���C�����j���[.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdMenu(11)"
      Tab(1).Control(1)=   "cmdMenu(9)"
      Tab(1).Control(2)=   "cmdMenu(6)"
      Tab(1).Control(3)=   "cmdMenu(2)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "�����e�i���X"
      TabPicture(2)   =   "���C�����j���[.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "cmdMenu(103)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdMenu(5)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdMenu(102)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdMenu(104)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdMenu(105)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdMenu(101)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdMenu(107)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "�V�X�e�����"
      TabPicture(3)   =   "���C�����j���[.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdMenu(106)"
      Tab(3).ControlCount=   1
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�ی�҃}�X�^���� �Ɖ�"
         Height          =   495
         Index           =   107
         Left            =   3480
         TabIndex        =   26
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�U�֗\��\�捞���ʃ��X�g"
         Height          =   495
         Index           =   11
         Left            =   -71520
         TabIndex        =   20
         Top             =   1140
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�����U�ֈ˗���(����)"
         Height          =   495
         Index           =   10
         Left            =   -74340
         TabIndex        =   19
         Top             =   540
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�U�ֈ˗���(�捞)"
         Height          =   495
         Index           =   9
         Left            =   -74340
         TabIndex        =   18
         Top             =   540
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "���Z�@�փf�[�^�捞"
         Height          =   495
         Index           =   6
         Left            =   -74340
         TabIndex        =   15
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�U�֗\��\ �� ���ʒm��(�捞)"
         Height          =   495
         Index           =   2
         Left            =   -74340
         TabIndex        =   14
         Top             =   1140
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�����U�փf�[�^�쐬(����)"
         Height          =   495
         Index           =   3
         Left            =   -74340
         TabIndex        =   13
         Top             =   2940
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�U�֗\��\ �� ���ʒm��(�ݐ�)"
         Height          =   495
         Index           =   4
         Left            =   -71520
         TabIndex        =   12
         Top             =   2940
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�����U�փf�[�^�쐬(�\��)"
         Height          =   495
         Index           =   1
         Left            =   -74340
         TabIndex        =   11
         Top             =   1740
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�����U�ֈ˗���(���)"
         Height          =   495
         Index           =   7
         Left            =   -74340
         TabIndex        =   10
         Top             =   1140
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�U�֗\��\(���)"
         Height          =   495
         Index           =   8
         Left            =   -74340
         Style           =   1  '���̨���
         TabIndex        =   9
         Top             =   2340
         Width           =   2295
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�ϑ��҃}�X�^�����e�i���X"
         Height          =   495
         Index           =   101
         Left            =   660
         TabIndex        =   8
         Top             =   540
         Width           =   2355
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�x���}�X�^�����e�i���X"
         Height          =   495
         Index           =   105
         Left            =   660
         TabIndex        =   7
         Top             =   3000
         Width           =   2355
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "���Z�@�փ}�X�^�����e�i���X"
         Height          =   495
         Index           =   104
         Left            =   660
         TabIndex        =   6
         Top             =   2400
         Width           =   2355
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�_��҃}�X�^�����e�i���X"
         Height          =   495
         Index           =   102
         Left            =   660
         TabIndex        =   5
         Top             =   1200
         Width           =   2355
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�_��҃}�X�^�f�[�^�쐬"
         Height          =   495
         Index           =   5
         Left            =   3480
         TabIndex        =   4
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "�ی�҃}�X�^�����e�i���X"
         Height          =   495
         Index           =   103
         Left            =   660
         TabIndex        =   3
         Top             =   1800
         Width           =   2355
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "��{���o�^"
         Height          =   495
         Index           =   106
         Left            =   -74340
         TabIndex        =   2
         Top             =   540
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "�I��(&X)"
      Height          =   495
      Index           =   0
      Left            =   5340
      TabIndex        =   0
      Top             =   3900
      Width           =   1335
   End
   Begin VB.Frame fraTimer 
      BorderStyle     =   0  '�Ȃ�
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   3495
      Begin VB.Label lblClientTime 
         Caption         =   "2007/06/13 13:58:11"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
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
            Name            =   "�l�r �o�S�V�b�N"
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
         Alignment       =   1  '�E����
         AutoSize        =   -1  'True
         Caption         =   "Client Time�F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
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
         Alignment       =   1  '�E����
         AutoSize        =   -1  'True
         Caption         =   "Server Time�F"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
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
         Name            =   "�l�r �o�S�V�b�N"
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
      TabIndex        =   27
      Top             =   4320
      Width           =   3555
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
    eFrmFurikaeIraishoInput = 10      '//eFrmHogoshaMaster
    eFrmFurikaeIraishoPrint = 7
    eFrmYoteiDataPrint = 8
    eFrmYoteiDataExport = 1
    eFrmYoteiReqImport = 2
    eFrmJissekiDataExport = 3
    eFrmJissekiDataRuiseki = 4
    '//Right Menu-2
    eFrmKeiyakushaMasterExport = 5
    eFrmBankDataImport = 6
    '//Right Menu
    eFrmItakushaMaster = 101
    eFrmKeiyakushaMaster = 102
    eFrmHogoshaMaster = 103
    eFrmBankMaster = 104
    eFrmHolidayMaster = 105
    eFrmSystemInfomation = 106
'//2006/03/02 �ی�҃}�X�^�捞�ǉ�
    eFrmFurikaeReqImport = 9
    eFrmYoteiReqImportReport = 11
    eFrmHogoshaMasterRireki = 107
End Enum

Private Sub cmdMenu_Click(Index As Integer)
    Dim frm As Form
    Select Case Index
    Case eButton.eEnd
        Unload Me       'Unload()�Ƀf�X�g���N�^����
    Case eButton.eFrmItakushaMaster
        Set frm = frmItakushaMaster
    Case eButton.eFrmHogoshaMaster, eButton.eFrmFurikaeIraishoInput         '//eFrmHogoshaMaster
        Set frm = frmHogoshaMaster
    Case eButton.eFrmYoteiReqImport
        Set frm = frmYoteiReqImport
    Case eButton.eFrmJissekiDataRuiseki
        Set frm = frmJissekiDataRuiseki
    Case eButton.eFrmKeiyakushaMaster
        Set frm = frmKeiyakushaMaster
    Case eButton.eFrmSystemInfomation
        Set frm = frmSystemInfomation
    Case eButton.eFrmHolidayMaster
        Set frm = frmHolidayMaster
    Case eButton.eFrmYoteiDataExport
        Set frm = frmYoteiDataExport
        frm.chkJisseki.Value = 0
    Case eButton.eFrmJissekiDataExport
        Set frm = frmYoteiDataExport
        frm.chkJisseki.Value = 1
    Case eButton.eFrmBankMaster
        Set frm = frmBankMaster
    Case eButton.eFrmKeiyakushaMasterExport
        Set frm = frmKeiyakushaMasterExport
    Case eButton.eFrmBankDataImport
        Set frm = frmBankDataImport
    Case eButton.eFrmFurikaeIraishoPrint
        Set frm = frmFurikaeIraishoPrint
    Case eButton.eFrmYoteiDataPrint
        Set frm = frmYoteiDataPrint
    Case eButton.eFrmFurikaeReqImport
        Set frm = frmFurikaeReqImport
    Case eButton.eFrmYoteiReqImportReport
        Set frm = frmYoteiReqImportReport
    Case eButton.eFrmHogoshaMasterRireki
        Set frm = frmHogoshaMasterRireki
    End Select
    '//�{�^�������������̂݋L������
    mReg.MenuButton = Index
    mReg.MenuTab = tabMenu.Tab
    If UCase(TypeName(frm)) <> UCase("Nothing") Then
        Set gdForm = Me
        Call frm.Show
        Call Me.Hide
    End If
End Sub

Private Sub Form_Activate()
    '//SetFocus �o���Ȃ����̃G���[�Ή�
    On Error Resume Next
    Call cmdMenu(mReg.MenuButton).SetFocus
End Sub

Private Sub Form_Load()
    Call mForm.Init(Me, gdDBS)
    cmdMenu(eButton.eFrmYoteiReqImport).Caption = " �U�֋��z�\��\" & vbCrLf & "�� ���ʒm�� (�捞)"
    cmdMenu(eButton.eFrmYoteiReqImportReport).Caption = " �U�֋��z�\��\" & vbCrLf & " �捞���ʃ��X�g"
    cmdMenu(eButton.eFrmJissekiDataRuiseki).Caption = " �U�֋��z�\��\" & vbCrLf & "�� ���ʒm�� (�ݐ�)"
    Call mForm.MoveSysDate
    tabMenu.Tab = mReg.MenuTab
    
    tmrTimer.Interval = 60000    '// �P�b��1,000 / �P����60,000
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
    '//SetFocus �o���Ȃ����̃G���[�Ή�
    On Error Resume Next
    Call cmdMenu(mReg.MenuButton).SetFocus
End Sub

Private Sub tmrTimer_Timer()
    lblClientTime.Caption = Format(Now(), "yyyy/MM/dd HH:nn:ss")
    lblServerTime.Caption = gdDBS.sysDate("yyyy/mm/dd hh24:mi:ss")
End Sub
