VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmFurikaeYoteiImport 
   Caption         =   "�U�֗\��\ �� ���ʒm��(�捞)"
   ClientHeight    =   7515
   ClientLeft      =   1125
   ClientTop       =   2400
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   11175
   Begin VB.Frame fraDetailInfo 
      Caption         =   "���׏W�v���"
      Height          =   1155
      Left            =   9120
      TabIndex        =   18
      Top             =   5160
      Width           =   1935
      Begin VB.Label Label2 
         Alignment       =   1  '�E����
         Caption         =   "�ύX�����F"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  '�E����
         Caption         =   "���z���v�F"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  '�E����
         Caption         =   "��񌏐��F"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblDetailCancel 
         Alignment       =   1  '�E����
         Caption         =   "123,456"
         BeginProperty Font 
            Name            =   "�l�r �o����"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblDetailCount 
         Alignment       =   1  '�E����
         Caption         =   "123,456"
         BeginProperty Font 
            Name            =   "�l�r �o����"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblDetailKingaku 
         Alignment       =   1  '�E����
         Caption         =   "123,456"
         BeginProperty Font 
            Name            =   "�l�r �o����"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   19
         Top             =   540
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSprUpdate 
      Caption         =   "�X�V(&S)"
      Height          =   435
      Left            =   9300
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ComboBox cboFIITKB 
      BackColor       =   &H000000FF&
      Height          =   300
      ItemData        =   "�U�֗\��\�捞.frx":0000
      Left            =   6240
      List            =   "�U�֗\��\�捞.frx":000D
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   17
      Top             =   60
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "�`�F�b�N(&C)"
      Height          =   435
      Left            =   2040
      TabIndex        =   7
      Top             =   6540
      Width           =   1395
   End
   Begin VB.CommandButton cmdErrList 
      Caption         =   "�G���[���X�g(&P)"
      Height          =   435
      Left            =   3540
      TabIndex        =   8
      Top             =   6540
      Width           =   1395
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�}�X�^���f(&U)"
      Height          =   435
      Left            =   7980
      TabIndex        =   10
      Top             =   6540
      Width           =   1395
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "�p��(&D)"
      Height          =   435
      Left            =   6480
      TabIndex        =   9
      Top             =   6540
      Width           =   1395
   End
   Begin VB.ComboBox cboImpDate 
      Height          =   300
      Left            =   1200
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   1
      Top             =   60
      Width           =   1935
   End
   Begin VB.ComboBox cboSort 
      Height          =   300
      ItemData        =   "�U�֗\��\�捞.frx":0037
      Left            =   4500
      List            =   "�U�֗\��\�捞.frx":0041
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   2
      Top             =   60
      Width           =   1335
   End
   Begin VB.Frame fraProgressBar 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "fraProgressBar"
      ForeColor       =   &H80000004&
      Height          =   290
      Left            =   1980
      TabIndex        =   13
      Top             =   7140
      Width           =   7060
      Begin MSComctlLib.ProgressBar pgrProgressBar 
         Height          =   255
         Left            =   15
         TabIndex        =   14
         Top             =   15
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
      End
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "�捞(&I)"
      Height          =   435
      Left            =   420
      TabIndex        =   6
      Top             =   6540
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "�I��(&X)"
      Height          =   435
      Left            =   9600
      TabIndex        =   0
      Top             =   6540
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  '������
      Height          =   315
      Left            =   0
      TabIndex        =   12
      Top             =   7200
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "�c�� 9,999 ��"
            TextSave        =   "�c�� 9,999 ��"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o����"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   10560
      Top             =   3540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ORADCLibCtl.ORADC dbcImportTotal 
      Height          =   315
      Left            =   9180
      Top             =   4080
      Visible         =   0   'False
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
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
      RecordSource    =   "select * from tfFurikaeYoteiImport Where firkbn=1"
   End
   Begin ORADCLibCtl.ORADC dbcImportDetail 
      Height          =   315
      Left            =   9180
      Top             =   4440
      Visible         =   0   'False
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
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
      RecordSource    =   "select * from tfFurikaeYoteiImport Where firkbn=0"
   End
   Begin FPSpread.vaSpread sprTotal 
      Bindings        =   "�U�֗\��\�捞.frx":0057
      Height          =   2865
      Left            =   420
      TabIndex        =   3
      Top             =   480
      Width           =   10140
      _Version        =   196608
      _ExtentX        =   17886
      _ExtentY        =   5054
      _StockProps     =   64
      ButtonDrawMode  =   4
      DAutoCellTypes  =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o����"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   17
      MaxRows         =   12
      ScrollBars      =   2
      SpreadDesigner  =   "�U�֗\��\�捞.frx":0074
      UserResize      =   0
      VScrollSpecial  =   -1  'True
   End
   Begin FPSpread.vaSpread sprDetail 
      Bindings        =   "�U�֗\��\�捞.frx":07C0
      Height          =   2895
      Left            =   420
      TabIndex        =   4
      Top             =   3420
      Width           =   8610
      _Version        =   196608
      _ExtentX        =   15187
      _ExtentY        =   5106
      _StockProps     =   64
      ButtonDrawMode  =   4
      DAutoCellTypes  =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o����"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   18
      MaxRows         =   15
      ScrollBars      =   2
      SpreadDesigner  =   "�U�֗\��\�捞.frx":07DE
      UserResize      =   0
      VirtualScrollBuffer=   -1  'True
      VScrollSpecial  =   -1  'True
   End
   Begin VB.Label Label8 
      Caption         =   "�捞����"
      Height          =   180
      Left            =   360
      TabIndex        =   16
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "�\����"
      Height          =   180
      Left            =   3780
      TabIndex        =   15
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   255
      Left            =   8460
      TabIndex        =   11
      Top             =   0
      Width           =   1395
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
   Begin VB.Menu mnuSpread 
      Caption         =   "�X�v���b�h�ҏW(&S)"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuTitle 
         Caption         =   "�^�C�g��"
      End
      Begin VB.Menu mnuSprDelete 
         Caption         =   "���ׂ̍폜(&D)"
      End
      Begin VB.Menu mnuSprReset 
         Caption         =   "���ׂ̍폜������(&R)"
      End
   End
End
Attribute VB_Name = "frmFurikaeYoteiImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If NO_RELEASE Then
Option Explicit

#Const DETAIL_SEQN_ORDER = True         '//���ׂ͂r�d�p���ɁF���Ƃ̓ˍ��������ɂ����I
#Const BLOCK_CHECK = False              '//�`�F�b�N���̃u���b�N���������邩�H��\���F�f�o�b�N���̂�
#If BLOCK_CHECK = True Then             '//�`�F�b�N���̃u���b�N���������邩�H��\���F�f�o�b�N���̂�
Private mCheckBlocks As Integer
#End If

Private Type tpFurikaeTotal    '//���v���R�[�h
    MochikomiBi As String * 8   '//�������ݓ� 2006/03/24 ���ڒǉ�
    KeiyakuNo   As String * 5   '//�_��Ҕԍ�
    KyoshitsuNo As String * 3   '//�����ԍ�
    PageNumber  As String * 2   '//�y�[�W�ԍ�
    FurikaeDate As String * 8   '//�U�֓�
    DetailCnt   As String * 2   '//���׌���
    DetailGaku  As String * 6   '//���׍��v���z
    CancelCnt   As String * 2   '//���׉�񌏐�
    RecKubun    As String * 1   '//���t���O�𗬗p�F�P�����^�X�����v
    KouzaName   As String * 40  '//2006/04/26 �������`�l��
    CrLf        As String * 2  'CR + LF
End Type

Private Type tpFurikaeDetail   '//���׃��R�[�h
    MochikomiBi As String * 8   '//�������ݓ� 2006/03/24 ���ڒǉ�
    KeiyakuNo   As String * 5   '//�_��Ҕԍ�
    KyoshitsuNo As String * 3   '//�����ԍ�
    PageNumber  As String * 2   '//�y�[�W�ԍ�
    FurikaeDate As String * 8   '//�U�֓�
    HogoshaNo   As String * 4   '//�ی�Ҕԍ�
    HenkouGaku  As String * 6   '//�ύX���z
    CancelFlag  As String * 1   '//���t���O
    KouzaName   As String * 40  '//2006/04/26 �������`�l��
    CrLf        As String * 2  'CR + LF
End Type

Private Enum eSprTotal
    eErrorStts = 1  '   FIERROR �G���[���e�F�ُ�A����A�x��
    eMochikomiBi    '           ������
    eImportCnt      '           �捞��
    eItakuName      '           �ϑ��Җ�
    eKeiyakuCode    '   FIKYCD  �_���
    eKeiyakuName    '           �_��Җ�
    eKyoshitsuNo    '   FIKSCD  �����ԍ�
    ePageNumber     '   FIPGNO  ��
    eFirukaeDate    '   FIFKDT  �U�֓�
    eHenkoCount     '   FIHKCT  �ύX����
    eHenkoKingaku   '   FIHKCT  �ύX���z
    eCancelCount    '   FIKYCT  ��񌏐�
    '//�\�������͍����܂�
    eUseCols
    eImpDate = eUseCols 'FIINDT
    eImpSEQ         '   FISEQN
    eItakuCode      '   FIITKB  �ϑ���
    eErrorFlag      '//�C������ FIERROR �ɍX�V����ׂɐݒ�F�˗����Ƃ͎኱�������Ⴄ�̂�...�B
    eEditFlag       '//�ύX�t���O
    eMaxCols = 30   '//�G���[����܂߂āI
End Enum
Private Enum eSprDetail
    eErrorStts = 1  '   FIERROR �G���[���e�F�ُ�A����A�x��
    eMochikomiBi    '           ������
    eImportCnt      '           �捞��
    eHogoshaNo      '   FIHGCD  �ی�Ҕԍ�
    eMasterKouza
    eImportKouza    '   FIKZNM  �������`�l��
    eHenkoGaku      '   FIHKKG  �ύX���z
    eCancelFlag     '   FIKYFG  ���t���O
    '//�\�������͍����܂�
    eUseCols
    eImpDate = eUseCols 'FIINDT
    eImpSEQ         '   FISEQN
    eItakuCode      '   FIITKB  �ϑ���
    eKeiyakuCode    '   FIKYCD  �_���
    eKyoshitsuNo    '   FIKSCD  �����ԍ�
    ePageNumber     '   FIPGNO  ��
    eFirukaeDate    '   FIFKDT  �U�֓�
    eErrorFlag      '//�C������ FIERROR �ɍX�V����ׂɐݒ�F�˗����Ƃ͎኱�������Ⴄ�̂�...�B
    eEditFlag       '//�ύX�t���O
    eMaxCols = 30   '//�G���[����܂߂āI
End Enum
Private mCaption    As String
Private mAbort      As Boolean
Private mForm       As New FormClass
Private mReg        As New RegistryClass
Private mYimp       As New FurikaeSchImpClass
Private mSprTotal   As New SpreadClass
Private mSprDetail  As New SpreadClass
Private mLeaveCellEvents As Boolean     '//�N�����̂P��ڂ̂� LeaveCell �C�x���g���������Ȃ��̂Ő���

Private Const cBtnCancel As String = "���~(&A)"
Private Const cBtnImport As String = "�捞(&I)"
Private Const cBtnDelete As String = "�p��(&D)"
Private Const cBtnCheck  As String = "�`�F�b�N(&C)"
Private Const cBtnUpdate As String = "�}�X�^���f(&U)"
Private Const cBtnSprUpdate As String = "�X�V(&S)"
Private Const cImportToYotei  As String = "Y"   '//�\�蔽�f
Private Const cImportToDelete As String = "D"   '//�p��
Private Const cEditDataMsg  As String = "�C�� => �`�F�b�N���������ĉ������B"
Private Const cImportMsg    As String = "�捞 => �`�F�b�N���������ĉ������B"
Private Const cDeleteMsg    As String = "�폜 => �������\�ł��B"
Private Const cVisibleRows  As Long = 12
Private Const cInSQLString = "FIINDT,FIITKB,FIKYCD,FIKSCD,FIPGNO"
'//2006/06/16 �_��Ҕԍ������̃p���`�f�[�^�Ή�
Private Const cFIKYCD_BadStart As Long = 90001  '//�_��҃p���`�����̊J�n�ԍ�
Private Const cFIITKB_BadCode As String = "Z"
'//���׍폜�̕ϐ��ݒ�
Private mDeleteSeqNo As Long        '//�폜�Ώۂr�d�p-�m��

Private mDeleteMenu As Integer      '//�폜�A�N�V�����̃��j���[ -1=Delete,0=NonMenu,1=Reset
Private Enum ePopup
    eDelete = -1
    eNoMenu
    eReset
End Enum

'//�r�p�k���ʃZ�b�g�̃I�[�_�[�� ���� �C���A�G���[�A�x���A����̏�
'//2006/04/14 ORDER ���v�f�ʂ�ɂȂ��Ă��Ȃ�����
'//2006/06/16 ���׍폜 -4 �̑Ή�
Private Const cSQLOrderString = " DECODE(FIERROR,-4,-13,-2, -11, -1,-12, 1,-10 ,FIERROR) "

Private Enum eSort
    eImportSeq
    eKeiyakusha
'    eKinnyuKikan
End Enum

Private Sub cboImpDate_Click()
    If "" = Trim(cboImpDate.Text) Then
        '//�L�蓾�Ȃ�
        Exit Sub
    End If
    If True = pSpreadCheckAndUpdate(sprTotal.Tag = mSprTotal.RowEdit Or sprDetail.Tag = mSprDetail.RowEdit) Then
        Exit Sub
    End If
    Dim ms As New MouseClass
    Call ms.Start
    '//�f�[�^�ǂݍ��݁� Spread �ɐݒ蔽�f
    Call pReadTotalDataAndSetting
End Sub

Private Sub cboSort_Click()
    Call cboImpDate_Click
End Sub

Private Function pMoveTempRecords(vCondition As String, vMode As String) As Long
    Dim sql As String
    '//�폜�Ώۃf�[�^�� Temp �Ƀo�b�N�A�b�v
    sql = "INSERT INTO " & mYimp.TfFurikaeImport & "Temp" & vbCrLf
    sql = sql & " SELECT SYSDATE,'" & vMode & "',a.*"
    sql = sql & " FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
    sql = sql & " WHERE 1 = 1" & vbCrLf
    sql = sql & vCondition
    Call gdDBS.Database.ExecuteSQL(sql)
    
    sql = "DELETE " & mYimp.TfFurikaeImport & " a " & vbCrLf
    sql = sql & " WHERE 1 = 1" & vbCrLf
    sql = sql & vCondition
    pMoveTempRecords = gdDBS.Database.ExecuteSQL(sql)
End Function

Private Function pProgressBarSet(ByRef rBlockStep As Integer, Optional ByRef rStepCnt As Long = -1) As Boolean
    DoEvents    '//�C�x���g��t
    If mAbort Then
        pProgressBarSet = False     '//�������f�I
        Exit Function
    End If
    '//�X�e�[�^�X�s�̐���E����
    If 0 <= rStepCnt Then
        If 0 = rStepCnt Then
            rBlockStep = rBlockStep - 1
        End If
        rStepCnt = rStepCnt + 1
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�c��(" & rBlockStep & ") - " & pgrProgressBar.Max - rStepCnt
        pgrProgressBar.Value = IIf(rStepCnt < pgrProgressBar.Max, rStepCnt, pgrProgressBar.Max)
    Else
        rBlockStep = rBlockStep - 1
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�c��(" & rBlockStep & ")"
        pgrProgressBar.Value = IIf(0 <= pgrProgressBar.Max - rBlockStep, pgrProgressBar.Max - rBlockStep, pgrProgressBar.Max)
    End If
    pProgressBarSet = True
#If BLOCK_CHECK = True Then           '//�`�F�b�N���̃u���b�N���������邩�H��\���F�f�o�b�N���̂�
    If rStepCnt <= 1 Then
        mCheckBlocks = mCheckBlocks + 1
    End If
#End If
End Function

Private Function pDataCheck(vImpDate As Variant) As Boolean
    Dim sqlStep As Long, Block As Integer, recCnt As Long
    
    Const cMaxBlock As Integer = 16
    Block = cMaxBlock
#If BLOCK_CHECK = True Then           '//�`�F�b�N���̃u���b�N���������邩�H��\���F�f�o�b�N���̂�
    mCheckBlocks = 0
#End If
    '// WHERE ��ɂ͕K���t��
    Dim SameConditions As String
    SameConditions = " AND FIINDT = TO_DATE('" & vImpDate & "','yyyy/mm/dd hh24:mi:ss')"
'//2006/06/16 ���׍폜�Ή�
    SameConditions = SameConditions & " AND FIERROR <> " & mYimp.errDeleted
    
    On Error GoTo gDataCheckError:
    
    Call gdDBS.AutoLogOut(mCaption, "[" & vImpDate & "] �̃`�F�b�N�������J�n����܂����B")
    
    Call gdDBS.Database.BeginTrans          '//�g�����U�N�V�����J�n
    Dim sql As String
    
    fraProgressBar.Visible = True
    pgrProgressBar.Max = cMaxBlock
    '//////////////////////////////////////////////////
    '//�G���[���ڃ��Z�b�g
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    '//�U�֗\��\�͍����ŃN���A���Ȃ��Ɖ����ł��o���Ȃ�
    sql = sql & " FIOKFG = " & mYimp.errNormal & "," & vbCrLf
    sql = sql & mYimp.StatusColumns(" = " & mYimp.errNormal & "," & vbCrLf)
    sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " FIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE 1 = 1" & vbCrLf  '//���܂��Ȃ�
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//�_��҃R�[�h�F�ϑ��҃R�[�h�����肷��ׂɐ�Ƀ`�F�b�N����
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    sql = sql & " FIKYCDE = DECODE(LENGTH(FIKYCD),5," & mYimp.errNormal & "," & mYimp.errInvalid & ")," & vbCrLf   '//�T���łȂ���΃G���[
    sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " FIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE FIKYCDE = " & mYimp.errNormal & vbCrLf
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//�ϑ��ҋ敪
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    sql = sql & " FIITKBE = (SELECT DECODE(COUNT(*),0," & mYimp.errInvalid & "," & mYimp.errNormal & ") " & vbCrLf
    sql = sql & "            FROM taItakushaMaster " & vbCrLf
    sql = sql & "            WHERE ABKYTP = SUBSTRB(a.FIKYCD,1,1)" & vbCrLf
    sql = sql & "            )," & vbCrLf
    sql = sql & " FIITKB  = (SELECT ABITKB "
    sql = sql & "            FROM taItakushaMaster " & vbCrLf
    sql = sql & "            WHERE ABKYTP = SUBSTRB(a.FIKYCD,1,1)" & vbCrLf
    sql = sql & "            )," & vbCrLf
    sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " FIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE FIITKBE = " & mYimp.errNormal & vbCrLf
    sql = sql & "   AND FIKYCDE = " & mYimp.errNormal & vbCrLf    '//��ł̌_��҃R�[�h�G���[�͕s�v
    sql = sql & "   AND FIITKB IS NULL" & vbCrLf        '//���Ɉϑ��ҋ敪�����͂���Ă���΍X�V���Ȃ�
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//�_��҃R�[�h�F�X�ɍēx�A�ϑ��Ҕz���̌_��҂��`�F�b�N
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    sql = sql & " FIKYCDE = (SELECT DECODE(COUNT(*),0," & mYimp.errInvalid & "," & mYimp.errNormal & ") " & vbCrLf
    sql = sql & "            FROM tbKeiyakushaMaster " & vbCrLf
    sql = sql & "            WHERE BAITKB = a.FIITKB " & vbCrLf
    sql = sql & "              AND BAKYCD = a.FIKYCD " & vbCrLf
    sql = sql & "              AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN BAKYST AND BAKYED " & vbCrLf '//�_�����
    sql = sql & "              AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN BAFKST AND BAFKED " & vbCrLf '//�U�֊���
    sql = sql & "         )," & vbCrLf
    sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " FIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE FIKYCD IS NOT NULL " & vbCrLf
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//�����ԍ�
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    '//JOIN�Ŏg�p���Ȃ��̂ŋ����ԍ������͂���Ă���݂̂𔻒f�ŉI
#If 1 Then
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    sql = sql & " FIKSCDE = DECODE(FIKSCD,NULL," & mYimp.errInvalid & "," & mYimp.errNormal & ")," & vbCrLf
    sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " FIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE FIKSCDE = " & mYimp.errNormal & vbCrLf
    sql = sql & SameConditions & vbCrLf
#Else
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    sql = sql & " FIKSCDE = (SELECT DECODE(COUNT(*),0," & mYimp.errInvalid & "," & mYimp.errNormal & ") " & vbCrLf
    sql = sql & "            FROM tbKeiyakushaMaster " & vbCrLf
    sql = sql & "            WHERE BAITKB = a.FIITKB " & vbCrLf
    sql = sql & "              AND BAKYCD = a.FIKYCD " & vbCrLf
    sql = sql & "              AND BAKSCD = a.FIKSCD " & vbCrLf
    sql = sql & "              AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN BAKYST AND BAKYED " & vbCrLf '//�_�����
    sql = sql & "              AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN BAFKST AND BAFKED " & vbCrLf '//�U�֊���
    sql = sql & "            )," & vbCrLf
    sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " FIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE FIKSCDE = " & mYimp.errNormal & vbCrLf
    sql = sql & SameConditions & vbCrLf
#End If
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//�ی�҃R�[�h�F�L��
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    sql = sql & " FIHGCDE = " & mYimp.errInvalid & "," & vbCrLf
    sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " FIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE FIHGCD IS NULL" & vbCrLf
    sql = sql & "   AND FIHGCDE = " & mYimp.errNormal & vbCrLf
    sql = sql & "   AND FIRKBN <> " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//�ی�҃R�[�h�F�ی�҃}�X�^
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    sql = sql & " FIHGCDE = (SELECT DECODE(COUNT(*),0," & mYimp.errInvalid & "," & mYimp.errNormal & ") " & vbCrLf
    sql = sql & "            FROM tcHogoshaMaster " & vbCrLf
    sql = sql & "            WHERE CAITKB = a.FIITKB " & vbCrLf
    sql = sql & "              AND CAKYCD = a.FIKYCD " & vbCrLf
    sql = sql & "              AND CAKSCD = a.FIKSCD " & vbCrLf    '//2006/04/13 �����ǉ�
    sql = sql & "              AND CAHGCD = a.FIHGCD " & vbCrLf
#If 0 Then  '//2006/04/05 ���݂���΃G���[�ɂ��Ȃ�
    '//�ی�҂͌��ݗL�����F�_����ԁ��U�֊���
'    sql = sql & "              AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN CAKYST AND CAKYED " & vbCrLf '//�_�����
'    sql = sql & "              AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN CAFKST AND CAFKED " & vbCrLf '//�U�֊���
#End If
    sql = sql & "            )," & vbCrLf
    sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " FIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE FIHGCDE = " & mYimp.errNormal & vbCrLf
    sql = sql & "   AND FIRKBN <> " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//�������`�l���F�ی�҃}�X�^
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    sql = sql & " FIKZNME = (SELECT " & vbCrLf
    sql = sql & "             CASE WHEN REPLACE(FIKZNM,' ',NULL) = REPLACE(CAKZNM,' ',NULL) THEN " & mYimp.errNormal & vbCrLf
    sql = sql & "                  ELSE                                                          " & mYimp.errInvalid & vbCrLf
    sql = sql & "             END " & vbCrLf
    sql = sql & "            FROM tcHogoshaMaster " & vbCrLf
    sql = sql & "            WHERE    (CAITKB,CAKYCD,CAKSCD,CAHGCD,    CASQNO) IN (" & vbCrLf
    sql = sql & "               SELECT CAITKB,CAKYCD,CAKSCD,CAHGCD,MAX(CASQNO)" & vbCrLf
    sql = sql & "               FROM tcHogoshaMaster " & vbCrLf
    sql = sql & "               WHERE CAITKB = a.FIITKB " & vbCrLf
    sql = sql & "                 AND CAKYCD = a.FIKYCD " & vbCrLf
    sql = sql & "                 AND CAKSCD = a.FIKSCD " & vbCrLf    '//2006/04/13 �����ǉ�
    sql = sql & "                 AND CAHGCD = a.FIHGCD " & vbCrLf
    sql = sql & "               GROUP BY CAITKB,CAKYCD,CAKSCD,CAHGCD" & vbCrLf
    sql = sql & "               )" & vbCrLf
    sql = sql & "              AND CAITKB = a.FIITKB " & vbCrLf
    sql = sql & "              AND CAKYCD = a.FIKYCD " & vbCrLf
    sql = sql & "              AND CAKSCD = a.FIKSCD " & vbCrLf    '//2006/04/13 �����ǉ�
    sql = sql & "              AND CAHGCD = a.FIHGCD " & vbCrLf
#If 0 Then  '//2006/04/05 ���݂���΃G���[�ɂ��Ȃ�
    '//�ی�҂͌��ݗL�����F�_����ԁ��U�֊���
'    sql = sql & "              AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN CAKYST AND CAKYED " & vbCrLf '//�_�����
'    sql = sql & "              AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN CAFKST AND CAFKED " & vbCrLf '//�U�֊���
#End If
    sql = sql & "            )," & vbCrLf
    sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " FIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE FIHGCDE = " & mYimp.errNormal & vbCrLf
    sql = sql & "   AND FIRKBN <> " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//�ی�҃R�[�h�F�U�֗\��f�[�^
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    sql = sql & " FIFKDTE = (SELECT DECODE(COUNT(*),0," & mYimp.errInvalid & "," & mYimp.errNormal & ") " & vbCrLf
    sql = sql & "            FROM tfFurikaeYoteiData " & vbCrLf
    sql = sql & "            WHERE FAITKB = a.FIITKB " & vbCrLf
    sql = sql & "              AND FAKYCD = a.FIKYCD " & vbCrLf
    sql = sql & "              AND FAHGCD = a.FIHGCD " & vbCrLf
    sql = sql & "              AND FASQNO = a.FIFKDT " & vbCrLf
    sql = sql & "            )," & vbCrLf
    sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " FIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE FIFKDTE = " & mYimp.errNormal & vbCrLf
    sql = sql & "   AND FIRKBN <> " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//�ی�҃R�[�h�F�U�֗\��f�[�^�F���v�ւ̓]�L
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    sql = sql & " FIFKDTE = (SELECT DECODE(COUNT(*),0," & mYimp.errNormal & "," & mYimp.errWarning & ") " & vbCrLf
    sql = sql & "            FROM " & mYimp.TfFurikaeImport & " b " & vbCrLf
    sql = sql & "            WHERE a.FIINDT = b.FIINDT " & vbCrLf
    sql = sql & "              AND a.FISEQN = b.FIRKBN " & vbCrLf
    sql = sql & "              AND b.FIRKBN <> " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & "              AND b.FIFKDTE <> " & mYimp.errNormal & vbCrLf
    sql = sql & "            )," & vbCrLf
    sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " FIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE FIFKDTE = " & mYimp.errNormal & vbCrLf
    sql = sql & "   AND FIRKBN  = " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//��񁕋��z�L��̃`�F�b�N
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    sql = sql & " FIHKKGE = " & mYimp.errWarning & "," & vbCrLf
    sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " FIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE FIHKKGE = " & mYimp.errNormal & vbCrLf
    '//2006/04/05 ���ŋ��z���L��F�x��
    sql = sql & "   AND ( NVL(FIKYFG,0) <> 0 AND NVL(FIHKKG,0) <> 0" & vbCrLf
    '//2006/04/05 ���Ŗ������z���Ȃ��F�x��
    '//2006/04/13 ���z�u�O�v�ŉ��Ŗ����f�[�^�L��
    'sql = sql & "      OR NVL(FIKYFG,0) =  0 AND NVL(FIHKKG,0)  = 0" & vbCrLf
    sql = sql & "   ) " & vbCrLf
    sql = sql & "   AND FIRKBN <> " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//���׍s�����v�s �Ԃ̃`�F�b�N
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
#If ORA_DEBUG = 1 Then
    Dim dynM As OraDynaset, dynS As OraDynaset, hkctErr As Boolean, hkkgErr As Boolean, kyctErr As Boolean
#Else
    Dim dynM As Object, dynS As Object, hkctErr As Boolean, hkkgErr As Boolean, kyctErr As Boolean
#End If
    '//���v���R�[�h�̎擾
    sql = "SELECT * FROM " & mYimp.TfFurikaeImport & vbCrLf
    sql = sql & " WHERE FIRKBN = " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & SameConditions & vbCrLf
    sql = sql & " ORDER BY FIITKB,FIKYCD,FIKSCD" & vbCrLf
#If ORA_DEBUG = 1 Then
    Set dynM = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dynM = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    sqlStep = 0
    If Not dynM.EOF Then
        pgrProgressBar.Value = 0
        pgrProgressBar.Max = dynM.RecordCount
    End If
    Do Until dynM.EOF
        '//////////////////////////////////////////////////
        '// DoEvents �� pProgressBarSet() �̒��Ŏ��s����Ă���
        If False = pProgressBarSet(Block, sqlStep) Then
            GoTo gDataCheckError:
        End If
        '//���׍s�̍��v���擾
        '//���X�|���X�͒x�������H
        sql = "SELECT " & vbCrLf
        '//�ύX�����ɂ͉����܂܂Ȃ�
        'sql = sql & " COUNT(*) FIHKCT,"& vbCrLf
        '//2006/04/14 ���z�u�O�v�ŉ�񖳂�������
        sql = sql & " SUM(DECODE(NVL(FIKYFG,0),0,1,0)) FIHKCT," & vbCrLf
        sql = sql & " SUM(       NVL(FIHKKG,0)       ) FIHKKG," & vbCrLf
        sql = sql & " SUM(DECODE(NVL(FIKYFG,0),0,0,1)) FIKYCT " & vbCrLf
        sql = sql & " FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
        sql = sql & " WHERE       (" & cInSQLString & ") IN(" & vbCrLf
        sql = sql & "       SELECT " & cInSQLString & vbCrLf
        sql = sql & "       FROM " & mYimp.TfFurikaeImport & " b " & vbCrLf
        sql = sql & "       WHERE FIINDT = TO_DATE(" & gdDBS.ColumnDataSet(dynM.Fields("FIINDT"), "D", vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss')" & vbCrLf
        sql = sql & "         AND FISEQN = " & gdDBS.ColumnDataSet(dynM.Fields("FISEQN"), "L", vEnd:=True) & vbCrLf
        sql = sql & "         AND FIRKBN = " & mYimp.RecordIsTotal & vbCrLf
        sql = sql & "       )"
'z 2006/06/13 �d���f�[�^���̏C��
'z        sql = sql & "   AND FIRKBN <> " & mYimp.RecordIsTotal & vbCrLf
        sql = sql & "   AND FISEQN <> " & gdDBS.ColumnDataSet(dynM.Fields("FISEQN"), "L", vEnd:=True) & vbCrLf
#If ORA_DEBUG = 1 Then
        Set dynS = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
        Set dynS = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
        '//2006/04/13 Null �G���[���:gdDBS.Nz()�ǉ�
        hkctErr = gdDBS.Nz(dynM.Fields("FIHKCT")) <> gdDBS.Nz(dynS.Fields("FIHKCT"))
        hkkgErr = gdDBS.Nz(dynM.Fields("FIHKKG")) <> gdDBS.Nz(dynS.Fields("FIHKKG"))
        kyctErr = gdDBS.Nz(dynM.Fields("FIKYCT")) <> gdDBS.Nz(dynS.Fields("FIKYCT"))
        '//���v�s�ɍX�V
        sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
        sql = sql & " FIHKCTE = " & IIf(hkctErr, mYimp.errWarning, mYimp.errNormal) & "," & vbCrLf
        sql = sql & " FIHKKGE = " & IIf(hkkgErr, mYimp.errWarning, mYimp.errNormal) & "," & vbCrLf
        sql = sql & " FIKYCTE = " & IIf(kyctErr, mYimp.errWarning, mYimp.errNormal) & "," & vbCrLf
        sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
        sql = sql & " FIUPDT = SYSDATE" & vbCrLf
        sql = sql & " WHERE FIINDT = TO_DATE(" & gdDBS.ColumnDataSet(dynM.Fields("FIINDT"), "D", vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss')" & vbCrLf
        sql = sql & "   AND FISEQN = " & gdDBS.ColumnDataSet(dynM.Fields("FISEQN"), "L", vEnd:=True) & vbCrLf
        sql = sql & "   AND FIRKBN = " & mYimp.RecordIsTotal & vbCrLf
        recCnt = gdDBS.Database.ExecuteSQL(sql)
        Call dynM.MoveNext
    Loop
    Call dynM.Close
    Set dynM = Nothing
    pgrProgressBar.Max = cMaxBlock
    '//////////////////////////////////////////////////
    '//�S�̃G���[���ڃZ�b�g�F�ŏ��ɐ���ɂ��Ă���̂Łu����v�t���O�͕s�v
    '//�ُ�f�[�^
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    sql = sql & " FIOKFG =  " & mYimp.updInvalid & "," & vbCrLf    '//�}�X�^���f�s��
    sql = sql & " FIERROR = " & mYimp.errInvalid & "," & vbCrLf
    sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " FIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE(" & vbCrLf
    sql = sql & mYimp.StatusColumns(" = " & mYimp.errInvalid & vbCrLf & " OR ", Len(vbCrLf & " OR ")) & vbCrLf & ")" & vbCrLf
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//�S�̃G���[���ڃZ�b�g�F�ŏ��ɐ���ɂ��Ă���̂Łu����v�t���O�͕s�v
    '//�x���f�[�^�F�}�X�^���f���Ȃ��f�[�^
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    sql = sql & " FIOKFG =  " & mYimp.updWarnErr & "," & vbCrLf   '//�}�X�^���f���Ȃ��t���O
    sql = sql & " FIERROR = " & mYimp.errWarning & "," & vbCrLf
    sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " FIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE FIERROR = " & mYimp.errNormal & vbCrLf    '//�ُ�Ŗ���
    sql = sql & "   AND FIOKFG <= " & mYimp.updNormal & vbCrLf
    sql = sql & "   AND(" & vbCrLf
    sql = sql & mYimp.StatusColumns(" >= " & mYimp.errWarning & vbCrLf & " OR ", Len(vbCrLf & " OR ")) & vbCrLf & ")" & vbCrLf
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//���ׂ̃G���[�����v�ɓ]�L
    '//////////////////////////////////////////////////
    Dim okFlag As Integer, erFlag As Integer
    '//���v���R�[�h�̎擾
    sql = "SELECT * FROM " & mYimp.TfFurikaeImport & vbCrLf
    sql = sql & " WHERE FIRKBN = " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & SameConditions & vbCrLf
    sql = sql & " ORDER BY FIITKB,FIKYCD,FIKSCD" & vbCrLf
#If ORA_DEBUG = 1 Then
    Set dynM = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dynM = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    sqlStep = 0
    If Not dynM.EOF Then
        pgrProgressBar.Value = 0
        pgrProgressBar.Max = dynM.RecordCount
    End If
    Do Until dynM.EOF
        '//////////////////////////////////////////////////
        '// DoEvents �� pProgressBarSet() �̒��Ŏ��s����Ă���
        If False = pProgressBarSet(Block, sqlStep) Then
            GoTo gDataCheckError:
        End If
        '//���׍s�̌��ʂ��擾
        '//���X�|���X�͒x�������H
        '//2006/04/13 NULL �G���[���
        sql = "SELECT NVL(MIN(NVL(FIOKFG,0)),0)  FIOKFG," & vbCrLf
        '//               ??ERROR => -1:�ُ�f�[�^ / 0:����f�[�^ / 1:�x���f�[�^ �ƂȂ��Ă���̂Œ���
        sql = sql & " NVL(MIN(NVL(FIERROR,0)),0) minERROR," & vbCrLf
        sql = sql & " NVL(MAX(NVL(FIERROR,0)),0) maxERROR " & vbCrLf
        sql = sql & " FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
        sql = sql & " WHERE       (" & cInSQLString & ") IN(" & vbCrLf
        sql = sql & "       SELECT " & cInSQLString & vbCrLf
        sql = sql & "       FROM " & mYimp.TfFurikaeImport & " b " & vbCrLf
        sql = sql & "       WHERE FIINDT = TO_DATE(" & gdDBS.ColumnDataSet(dynM.Fields("FIINDT"), "D", vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss')" & vbCrLf
        sql = sql & "         AND FISEQN = " & gdDBS.ColumnDataSet(dynM.Fields("FISEQN"), "L", vEnd:=True) & vbCrLf
        sql = sql & "         AND FIRKBN = " & mYimp.RecordIsTotal & vbCrLf
        sql = sql & "       )"
        sql = sql & "   AND FIRKBN <> " & mYimp.RecordIsTotal & vbCrLf
#If ORA_DEBUG = 1 Then
        Set dynS = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
        Set dynS = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
        '//���ׂ̃G���[�����v�����[���Ȃ�I�t���O�͕����Ȃ̂ŋt�]����
        '// OKFG �̔��f
        If Val(dynM.Fields("FIOKFG")) = mYimp.errNormal And Val(dynS.Fields("FIOKFG")) <> mYimp.errNormal Then
            okFlag = dynS.Fields("FIOKFG")
        Else
            okFlag = dynM.Fields("FIOKFG")
        End If
        '// ERROR �̔��f
        If Val(dynS.Fields("minERROR")) = mYimp.errNormal And Val(dynS.Fields("maxERROR")) = mYimp.errNormal Then
            '//���ׂ����ׂāu����v�Ȃ獇�v�̃G���[���
            erFlag = dynM.Fields("FIERROR")
        ElseIf Val(dynM.Fields("FIERROR")) = mYimp.errNormal Then
            '//���v���u����v
            If Val(dynS.Fields("minERROR")) = mYimp.errInvalid Then
                '//���ׂ��u�ُ�v�Ȃ獇�v���u�ُ�v
                erFlag = dynS.Fields("minERROR")
            ElseIf Val(dynS.Fields("maxERROR")) = mYimp.errWarning Then
                '//���ׂ��u�x���v�Ȃ獇�v���u�x���v
                erFlag = dynS.Fields("maxERROR")
            Else
                erFlag = dynM.Fields("FIERROR") '//���蓾�Ȃ��H
            End If
        ElseIf Val(dynM.Fields("FIERROR")) = mYimp.errWarning And Val(dynS.Fields("minERROR")) = mYimp.errInvalid Then
            '//���v���u�x���v�Ŗ��ׂ��u�ُ�v�Ȃ獇�v�́u�ُ�v
            erFlag = dynS.Fields("minERROR")
        Else
            erFlag = dynM.Fields("FIERROR")
        End If
        '//���v�s�ɍX�V
        sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
        sql = sql & " FIOKFG  = " & okFlag & "," & vbCrLf
        sql = sql & " FIERROR = " & erFlag & "," & vbCrLf
        sql = sql & " FIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
        sql = sql & " FIUPDT = SYSDATE" & vbCrLf
        sql = sql & " WHERE FIINDT = TO_DATE(" & gdDBS.ColumnDataSet(dynM.Fields("FIINDT"), "D", vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss')" & vbCrLf
        sql = sql & "   AND FISEQN = " & gdDBS.ColumnDataSet(dynM.Fields("FISEQN"), "L", vEnd:=True) & vbCrLf
        sql = sql & "   AND FIRKBN = " & mYimp.RecordIsTotal & vbCrLf
        recCnt = gdDBS.Database.ExecuteSQL(sql)
        Call dynM.MoveNext
    Loop
    Call dynM.Close
    Set dynM = Nothing
    pgrProgressBar.Max = cMaxBlock
    
    Call gdDBS.Database.CommitTrans         '//�g�����U�N�V��������I��
    fraProgressBar.Visible = False
    Call gdDBS.AutoLogOut(mCaption, "[" & vImpDate & "] �̃`�F�b�N�������������܂����B")
    '//�X�e�[�^�X�s�̐���E����
    stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�`�F�b�N����"
    pDataCheck = True

#If BLOCK_CHECK = True Then           '//�`�F�b�N���̃u���b�N���������邩�H��\���F�f�o�b�N���̂�
     Call MsgBox("�`�F�b�N�����u���b�N�� " & mCheckBlocks & " �ӏ��ł����B")
#End If
    
    Exit Function
gDataCheckError:
    fraProgressBar.Visible = False
    Call gdDBS.Database.Rollback            '//�g�����U�N�V�����ُ�I��
    If Err Then
        Dim errCode As Integer, errMsg As String
        If gdDBS.Database.LastServerErr Then
            errCode = gdDBS.Database.LastServerErr
            errMsg = gdDBS.Database.LastServerErrText
        Else
            errCode = Err
            errMsg = Error
        End If
        fraProgressBar.Visible = False
        '//�X�e�[�^�X�s�̐���E����
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�`�F�b�N�G���[(" & errCode & ")"
        Call gdDBS.AutoLogOut(mCaption, "[" & vImpDate & "] �̃`�F�b�N�������ɃG���[���������܂����B(Error=" & errCode & ")")
        Call MsgBox("�`�F�b�N�Ώ� = [" & cboImpDate.Text & "]" & vbCrLf & _
                    "�̓G���[�������������߃`�F�b�N�͒��~����܂����B" & vbCrLf & errMsg, _
                vbOKOnly + vbCritical, mCaption)
    Else
        Call gdDBS.AutoLogOut(mCaption, "[" & vImpDate & "] �̃`�F�b�N���������f����܂����B")
        '//�X�e�[�^�X�s�̐���E����
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�`�F�b�N���f"
    End If
End Function

Private Sub cmdCheck_Click()
    If True = pSpreadCheckAndUpdate(sprTotal.Tag = mSprTotal.RowEdit Or sprDetail.Tag = mSprDetail.RowEdit) Then
        Exit Sub
    End If
    If -1 <> pAbortButton(cmdCheck, cBtnCheck) Then
        Exit Sub
    End If
    cmdCheck.Caption = cBtnCancel
    '//�R�}���h�E�{�^������
    Call pLockedControl(False, cmdCheck)
    '//�`�F�b�N����
    If True = pDataCheck(cboImpDate.Text) Then
        '//�f�[�^�ǂݍ��݁� Spread �ɐݒ蔽�f
        Call pReadTotalDataAndSetting
    End If
    '//�{�^����߂�
    cmdCheck.Caption = cBtnCheck
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
End Sub

Private Sub cmdDelete_Click()
    If True = pSpreadCheckAndUpdate(sprTotal.Tag = mSprTotal.RowEdit Or sprDetail.Tag = mSprDetail.RowEdit) Then
        Exit Sub
    End If
    If vbOK <> MsgBox("���ݕ\������Ă���f�[�^��j�����܂�." & vbCrLf & vbCrLf & _
                      "�p���Ώ� = [" & cboImpDate.Text & "]" & vbCrLf & vbCrLf & _
                      "��낵���ł����H", vbOKCancel + vbInformation, mCaption) Then
        Exit Sub
    End If
    If -1 <> pAbortButton(cmdDelete, cBtnDelete) Then
        Exit Sub
    End If
    cmdDelete.Caption = cBtnCancel
    '//�R�}���h�E�{�^������
    Call pLockedControl(False, cmdDelete)
    
    Dim ms As New MouseClass, recCnt As Long
    Call ms.Start
    
    Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] �̔p�����J�n����܂����B")
    
    On Error GoTo cmdDelete_ClickErr:
    Call gdDBS.Database.BeginTrans
    
    '//�}�X�^���f���ɂ�������������̂ŋ��ʉ�
    recCnt = pMoveTempRecords(" AND FIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss')", cImportToDelete)
    If recCnt < 0 Then
        GoTo cmdDelete_ClickErr:
    End If
    
    Call gdDBS.Database.CommitTrans
    
    Set ms = Nothing
    Call MsgBox("�p���Ώ� = [" & cboImpDate.Text & "]" & vbCrLf & vbCrLf & _
                recCnt & " �����p������܂���.", vbOKOnly + vbInformation, mCaption)
    
    '//�X�e�[�^�X�s�̐���E����
    stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�p������"
    Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] �� " & recCnt & " ���̔p�����������܂����B")
    
    Call pMakeComboBox
    '//�{�^����߂�
    cmdDelete.Caption = cBtnDelete
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
    Exit Sub
cmdDelete_ClickErr:
    Call gdDBS.Database.Rollback
    If Err Then
        Dim errCode As Integer, errMsg As String
        If gdDBS.Database.LastServerErr Then
            errCode = gdDBS.Database.LastServerErr
            errMsg = gdDBS.Database.LastServerErrText
        Else
            errCode = Err
            errMsg = Error
        End If
        fraProgressBar.Visible = False
        '//�X�e�[�^�X�s�̐���E����
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�p���G���[(" & errCode & ")"
        Call gdDBS.AutoLogOut(mCaption, "�G���[�������������ߔp���͒��~����܂����B(Error=" & errMsg & ")")
        Call MsgBox("�p���Ώ� = [" & cboImpDate.Text & "]" & vbCrLf & _
                    "�̓G���[�������������ߔp���͒��~����܂����B" & vbCrLf & errMsg, _
                vbOKOnly + vbCritical, mCaption)
    Else
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�p�����f"
        Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] �̔p���͒��~����܂����B")
    End If
    '//�{�^����߂�
    cmdDelete.Caption = cBtnDelete
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
End Sub

Private Sub cmdEnd_Click()
    If True = pSpreadCheckAndUpdate(sprTotal.Tag = mSprTotal.RowEdit Or sprDetail.Tag = mSprDetail.RowEdit) Then
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub pLockedControl(blMode As Boolean, Optional vButton As CommandButton = Nothing)
    cmdImport.Enabled = blMode
    cmdCheck.Enabled = blMode
    cmdErrList.Enabled = blMode
    cmdDelete.Enabled = blMode
    cmdUpdate.Enabled = blMode
    cmdSprUpdate.Enabled = False    '//��Ɏg�p�s�ASpread �C���t���O�Ɠ����Ɏg�p��
    cmdEnd.Enabled = blMode     '//�����r���ŏI������Ƃ��������Ȃ�̂ŏI�����E���I
    If Not vButton Is Nothing Then
        vButton.Enabled = True
    End If
End Sub

Private Function pAbortButton(vButton As CommandButton, vCaption As String) As Integer
    pAbortButton = -1   '// -1 = �����J�n
    mAbort = False
    If vButton.Caption <> cBtnCancel Then
        Exit Function
    End If
    pAbortButton = MsgBox(Left(vCaption, InStr(vCaption, "(") - 1) & "�𒆎~���܂����H", vbInformation + vbOKCancel, mCaption)
    If vbOK <> pAbortButton Then
        Exit Function   '//���~����߂��I
    End If
    vButton.Caption = vCaption
    mAbort = True
End Function

Private Sub pReadTotalDataAndSetting()
    
    dbcImportTotal.RecordSource = pMakeSQLReadDataTotal
    '//���ڕҏW����̂ŉ��z���[�h�ɂ��Ȃ�
    'sprMeisai.VirtualMode = False   '//��U���z���[�h����
    Call dbcImportTotal.Refresh
    sprTotal.VScrollSpecial = True
    sprTotal.VScrollSpecialType = 0
    sprTotal.MaxRows = dbcImportTotal.Recordset.RecordCount
    '//�Z���P�ʂɃG���[�ӏ����J���[�\��
    Call pSpreadTotalSetErrorStatus(True)
    '//ToolTip ��L���ɂ���ׂɋ����I�Ƀt�H�[�J�X���ڂ��FForm_Load()���Ȃ̂ŃG���[�ɂȂ�I
    On Error Resume Next
    Call sprTotal.SetFocus
    mLeaveCellEvents = False    '//�N�����̂P��ڂ̂� LeaveCell �C�x���g���������Ȃ��̂Ő���
    '//�����\���͂O���A���v�N���b�N���ɕ\�������悤��...�B
    dbcImportDetail.RecordSource = ""
    sprDetail.MaxRows = 0
    '//���v�S�̂̏C���t���O�����Z�b�g
    sprTotal.Tag = mSprTotal.RowNonEdit
End Sub

Private Sub pReadDetailDataAndSetting(vImpDate As String, vSeqNo As Long)
    
    dbcImportDetail.RecordSource = pMakeSQLReadDataDetail(vImpDate, vSeqNo)
    '//���ڕҏW����̂ŉ��z���[�h�ɂ��Ȃ�
    'sprMeisai.VirtualMode = False   '//��U���z���[�h����
    Call dbcImportDetail.Refresh
    sprDetail.VScrollSpecial = True
    sprDetail.VScrollSpecialType = 0
    sprDetail.MaxRows = dbcImportDetail.Recordset.RecordCount
    
    '//���ׂ̍��v��\��
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    sql = "SELECT " & vbCrLf
    '//�ύX�����ɂ͉����܂܂Ȃ�
    'sql = sql & " COUNT(*) FIHKCT,"& vbCrLf
    sql = sql & " SUM(DECODE(NVL(FIKYFG,0),0,1,0)) FIHKCT," & vbCrLf
    sql = sql & " SUM(       NVL(FIHKKG,0)       ) FIHKKG," & vbCrLf
    sql = sql & " SUM(DECODE(NVL(FIKYFG,0),0,0,1)) FIKYCT " & vbCrLf
    sql = sql & " FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
    sql = sql & " WHERE   (" & cInSQLString & ") IN(" & vbCrLf
    sql = sql & "   SELECT " & cInSQLString & vbCrLf
    sql = sql & "   FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
    sql = sql & "   WHERE FIINDT = TO_DATE('" & vImpDate & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf
    sql = sql & "     AND FISEQN = " & gdDBS.ColumnDataSet(vSeqNo, vEnd:=True) & vbCrLf
    sql = sql & "   )" & vbCrLf
'z 2006/06/13 �d���f�[�^���̏C��
'z    sql = sql & "   AND FIRKBN <> " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & "   AND FIRKBN = " & vSeqNo & vbCrLf
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    lblDetailCount.Caption = Format(dyn.Fields("FIHKCT"), "#,##0")
    lblDetailKingaku.Caption = Format(dyn.Fields("FIHKKG"), "#,##0")
    lblDetailCancel.Caption = Format(dyn.Fields("FIKYCT"), "#,##0")
    Call dyn.Close
    Set dyn = Nothing

    '//�Z���P�ʂɃG���[�ӏ����J���[�\��
    Call pSpreadDetailSetErrorStatus(vImpDate, vSeqNo)
    '//���בS�̂̏C���t���O�����Z�b�g
    sprDetail.Tag = mSprDetail.RowEdit
End Sub

Private Function pMakeSQLReadDataTotal() As String
    Dim sql As String
    
    sql = "SELECT * FROM(" & vbCrLf
    sql = sql & "SELECT " & vbCrLf
    'sql = sql & " CIERROR," & vbCrLf
#If SHORT_MSG Then
    sql = sql & " DECODE(FIERROR,-4,'�폜',-3,'�捞',-2,'�C��',-1,'�ُ�',0,'����',1,'�x��','��O') as CIERRNM," & vbCrLf
#Else
    sql = sql & " CASE WHEN FIERROR = -2 THEN " & gdDBS.ColumnDataSet(cEditDataMsg, vEnd:=True) & vbCrLf
    sql = sql & "      WHEN FIERROR = -3 THEN " & gdDBS.ColumnDataSet(cImportMsg, vEnd:=True) & vbCrLf
'//2006/06/16 ���׍폜�Ή�
    sql = sql & "      WHEN FIERROR = -4 THEN " & gdDBS.ColumnDataSet(cDeleteMsg, vEnd:=True) & vbCrLf
    sql = sql & "      WHEN FIERROR IN(-1,+0,+1) THEN " & vbCrLf
    sql = sql & "           DECODE(FIERROR," & vbCrLf
    sql = sql & "               -1,'�ُ�'," & vbCrLf
    sql = sql & "               +0,'����'," & vbCrLf
    sql = sql & "               +1,'�x��'," & vbCrLf
    sql = sql & "               NULL" & vbCrLf
    sql = sql & "           ) || ' => ' || " & vbCrLf
    sql = sql & "       DECODE(FIOKFG," & vbCrLf
    sql = sql & "               " & mYimp.updInvalid & ",'" & mYimp.mUpdateMessage(mYimp.updInvalid) & "'," & vbCrLf
    sql = sql & "               " & mYimp.updWarnErr & ",'" & mYimp.mUpdateMessage(mYimp.updWarnErr) & "'," & vbCrLf
    sql = sql & "               " & mYimp.updNormal & ",'" & mYimp.mUpdateMessage(mYimp.updNormal) & "'," & vbCrLf
    sql = sql & "               " & mYimp.updWarnUpd & ",'" & mYimp.mUpdateMessage(mYimp.updWarnUpd) & "'," & vbCrLf
    '//����ȃf�[�^�͖���
    'sql = sql & "               " & mYimp.updResetCancel & ",'" & mYimp.mUpdateMessage(mYimp.updResetCancel) & "'," & vbCrLf
    sql = sql & "               '�������ʂ�����ł��܂���B'" & vbCrLf
    sql = sql & "           )" & vbCrLf
    sql = sql & "      ELSE                             '��O => �������ʂ�����ł��܂���B'" & vbCrLf
    sql = sql & " END as FIERRNM," & vbCrLf
'//2006/04/26 �������E�񐔕\��
    sql = sql & " TO_CHAR(TO_DATE(FIMCDT,'YYYYMMDD'),'yyyy/mm/dd') FIMCDT," & vbCrLf
'    sql = sql & " CASE WHEN NVL(FIICNT,0) <= 1 THEN NULL ELSE FIICNT END FIICNT," & vbCrLf
    sql = sql & " FIICNT," & vbCrLf
#End If
    sql = sql & " (SELECT ABKJNM " & vbCrLf
    sql = sql & "  FROM taItakushaMaster" & vbCrLf
    sql = sql & "  WHERE ABITKB = a.FIITKB" & vbCrLf
    sql = sql & " ) as ABKJNM," & vbCrLf    '//�ʏ�̊O�������ł���Ƃ�₱�����̂�...(tcHogoshaImport Table �͑S���o�������I)
    sql = sql & " FIKYCD," & vbCrLf
'//2006/04/13 ���������ʂ�����̂ŃG���[�ɂȂ�Ή� DISTINCT
    sql = sql & " (SELECT MAX(BAKJNM) BAKJNM " & vbCrLf
    sql = sql & "  FROM tbKeiyakushaMaster " & vbCrLf
    sql = sql & "  WHERE BAITKB = a.FIITKB" & vbCrLf
    sql = sql & "    AND BAKYCD = a.FIKYCD" & vbCrLf
'//2006/05/17 �ŐV�̌_��҂�\������ו��� : If 0 Then => If 1 Then
#If 1 Then  '//2006/04/05 ���݂���΃G���[�ɂ��Ȃ�
    '//�_��҂͌��ݗL�����F�_����ԁ��U�֊���
    sql = sql & "    AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN BAKYST AND BAKYED" & vbCrLf
    sql = sql & "    AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN BAFKST AND BAFKED" & vbCrLf
#End If
    sql = sql & " ) as BAKJNM," & vbCrLf    '//�ʏ�̊O�������ł���Ƃ�₱�����̂�...(tcHogoshaImport Table �͑S���o�������I)
    sql = sql & " FIKSCD," & vbCrLf
    sql = sql & " FIPGNO," & vbCrLf
    sql = sql & " TO_CHAR(TO_DATE(FIFKDT,'YYYYMMDD'),'yyyy/mm/dd') FIFKDT," & vbCrLf
    sql = sql & " FIHKCT," & vbCrLf
    sql = sql & " FIHKKG," & vbCrLf
    sql = sql & " FIKYCT," & vbCrLf
    sql = sql & " TO_CHAR(FIINDT,'yyyy/mm/dd hh24:mi:ss') FIINDT," & vbCrLf
    sql = sql & " FISEQN," & vbCrLf
    sql = sql & " FIITKB," & vbCrLf
    sql = sql & " FIERROR," & vbCrLf
    sql = sql & mSprTotal.RowNonEdit & " AS EditFlag "
    sql = sql & " FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
    sql = sql & " WHERE FIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf
    sql = sql & "   AND FIRKBN = " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & " ORDER BY " & cSQLOrderString & vbCrLf    '�C���A�G���[�A�x���A����̏�
    '//�ȍ~�̂n�q�c�d�q��
    Select Case cboSort.ListIndex
    Case eSort.eImportSeq
        sql = sql & ",FIINDT,FISEQN" & vbCrLf
    Case eSort.eKeiyakusha
        sql = sql & ",FIITKB,FIKYCD,FIKSCD,FIPGNO,FIFKDT,FIHGCD,FISEQN" & vbCrLf
    Case Else
    End Select
    sql = sql & ")" & vbCrLf
    pMakeSQLReadDataTotal = sql
End Function

Private Function pMakeSQLReadDataDetail(vDate As String, vSeqNo As Long) As String
    Dim sql As String
    
    sql = "SELECT * FROM(" & vbCrLf
    sql = sql & "SELECT " & vbCrLf
    'sql = sql & " CIERROR," & vbCrLf
#If SHORT_MSG Then
    sql = sql & " DECODE(FIERROR,-3,'�捞',-2,'�C��',-1,'�ُ�',0,'����',1,'�x��','��O') as CIERRNM," & vbCrLf
#Else
    sql = sql & " CASE WHEN FIERROR = -2 THEN " & gdDBS.ColumnDataSet(cEditDataMsg, vEnd:=True) & vbCrLf
    sql = sql & "      WHEN FIERROR = -3 THEN " & gdDBS.ColumnDataSet(cImportMsg, vEnd:=True) & vbCrLf
'//2006/06/16 ���׍폜�Ή�
    sql = sql & "      WHEN FIERROR = -4 THEN " & gdDBS.ColumnDataSet(cDeleteMsg, vEnd:=True) & vbCrLf
    sql = sql & "      WHEN FIERROR IN(-1,+0,+1) THEN " & vbCrLf
    sql = sql & "           DECODE(FIERROR," & vbCrLf
    sql = sql & "               -1,'�ُ�'," & vbCrLf
    sql = sql & "               +0,'����'," & vbCrLf
    sql = sql & "               +1,'�x��'," & vbCrLf
    sql = sql & "               NULL" & vbCrLf
    sql = sql & "           ) || ' => ' || " & vbCrLf
    sql = sql & "       DECODE(FIOKFG," & vbCrLf
    sql = sql & "               " & mYimp.updInvalid & ",'" & mYimp.mUpdateMessage(mYimp.updInvalid) & "'," & vbCrLf
    sql = sql & "               " & mYimp.updWarnErr & ",'" & mYimp.mUpdateMessage(mYimp.updWarnErr) & "'," & vbCrLf
    sql = sql & "               " & mYimp.updNormal & ",'" & mYimp.mUpdateMessage(mYimp.updNormal) & "'," & vbCrLf
    sql = sql & "               " & mYimp.updWarnUpd & ",'" & mYimp.mUpdateMessage(mYimp.updWarnUpd) & "'," & vbCrLf
    '//����ȃf�[�^�͖���
    'sql = sql & "               " & mYimp.updResetCancel & ",'" & mYimp.mUpdateMessage(mYimp.updResetCancel) & "'," & vbCrLf
    sql = sql & "               '�������ʂ�����ł��܂���B'" & vbCrLf
    sql = sql & "           )" & vbCrLf
    sql = sql & "      ELSE                             '��O => �������ʂ�����ł��܂���B'" & vbCrLf
    sql = sql & " END as FIERRNM," & vbCrLf
'//2006/04/26 �������E�񐔕\��
    sql = sql & " TO_CHAR(TO_DATE(FIMCDT,'YYYYMMDD'),'yyyy/mm/dd') FIMCDT," & vbCrLf
'    sql = sql & " CASE WHEN NVL(FIICNT,0) <= 1 THEN NULL ELSE FIICNT END FIICNT," & vbCrLf
    sql = sql & " FIICNT," & vbCrLf
#End If
    sql = sql & " FIHGCD," & vbCrLf
'//2006/04/13 ���������ʂ�����̂ŃG���[�ɂȂ�Ή� DISTINCT
'//2006/04/27 �p���`�f�[�^�Ɍ������`�l����ǉ������וύX CAKJNM=>CAKZNM
    sql = sql & " (SELECT DISTINCT CAKZNM " & vbCrLf
    sql = sql & "  FROM tcHogoshaMaster " & vbCrLf
    sql = sql & "  WHERE CAITKB = a.FIITKB" & vbCrLf
    sql = sql & "    AND CAKYCD = a.FIKYCD" & vbCrLf
    sql = sql & "    AND CAKSCD = a.FIKSCD" & vbCrLf    '//2006/04/13 �����ǉ�
    sql = sql & "    AND CAHGCD = a.FIHGCD" & vbCrLf
#If 0 Then  '//2006/04/05 ���݂���΃G���[�ɂ��Ȃ�
    '//�ی�҂͌��ݗL�����F�_����ԁ��U�֊���
    sql = sql & "    AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN CAKYST AND CAKYED" & vbCrLf
    sql = sql & "    AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN CAFKST AND CAFKED" & vbCrLf
#End If
'//2006/04/27 �p���`�f�[�^�Ɍ������`�l����ǉ������וύX CAKJNM=>CAKZNM
    sql = sql & " ) as CAKZNM," & vbCrLf    '//�ʏ�̊O�������ł���Ƃ�₱�����̂�...(tcHogoshaImport Table �͑S���o�������I)
'//2006/04/27 �p���`�f�[�^�Ɍ������`�l����ǉ������וύX
#If 1 Then
    sql = sql & " FIKZNM,"
#Else
    '//2006/04/05 ���҂�\��
    '//    sql = sql & " (SELECT CAKNNM " & vbCrLf
    '//2006/04/13 ���������ʂ�����̂ŃG���[�ɂȂ�Ή� DISTINCT
        sql = sql & " (SELECT DISTINCT DECODE(NVL(CAKYFG,0),0,CAKNNM,'(���)')" & vbCrLf
        sql = sql & "  FROM tcHogoshaMaster " & vbCrLf
        sql = sql & "  WHERE CAITKB = a.FIITKB" & vbCrLf
        sql = sql & "    AND CAKYCD = a.FIKYCD" & vbCrLf
        sql = sql & "    AND CAKSCD = a.FIKSCD" & vbCrLf    '//2006/04/13 �����ǉ�
        sql = sql & "    AND CAHGCD = a.FIHGCD" & vbCrLf
    #If 0 Then  '//2006/04/05 ���݂���΃G���[�ɂ��Ȃ�
        '//�ی�҂͌��ݗL�����F�_����ԁ��U�֊���
        sql = sql & "    AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN CAKYST AND CAKYED" & vbCrLf
        sql = sql & "    AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN CAFKST AND CAFKED" & vbCrLf
    #End If
        sql = sql & " ) as CAKNNM," & vbCrLf    '//�ʏ�̊O�������ł���Ƃ�₱�����̂�...(tcHogoshaImport Table �͑S���o�������I)
#End If
    sql = sql & " FIHKKG," & vbCrLf
    sql = sql & " FIKYFG," & vbCrLf
    sql = sql & " TO_CHAR(FIINDT,'yyyy/mm/dd hh24:mi:ss') FIINDT," & vbCrLf
    sql = sql & " FISEQN," & vbCrLf
    sql = sql & " FIITKB," & vbCrLf
    sql = sql & " FIKYCD," & vbCrLf
    sql = sql & " FIKSCD," & vbCrLf
    sql = sql & " FIPGNO," & vbCrLf
    sql = sql & " TO_CHAR(TO_DATE(FIFKDT,'YYYYMMDD'),'yyyy/mm/dd') FIFKDT," & vbCrLf
    sql = sql & " FIERROR," & vbCrLf
    sql = sql & mSprDetail.RowNonEdit & " AS EditFlag " & vbCrLf
    sql = sql & " FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
    sql = sql & " WHERE   (" & cInSQLString & ") IN(" & vbCrLf
    sql = sql & "   SELECT " & cInSQLString & vbCrLf
    sql = sql & "   FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
    sql = sql & "   WHERE FIINDT = TO_DATE('" & vDate & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf
    sql = sql & "     AND FISEQN = " & gdDBS.ColumnDataSet(vSeqNo, vEnd:=True) & vbCrLf
    sql = sql & "   )" & vbCrLf
'z 2006/06/13 �d���f�[�^���̏C��
'z    sql = sql & "   AND FIRKBN <> " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & "   AND FIRKBN = " & vSeqNo & vbCrLf
'//���ׂ͂r�d�p���ɁF���Ƃ̓ˍ��������ɂ����I
#If DETAIL_SEQN_ORDER = True Then
    sql = sql & " ORDER BY FIINDT,FISEQN" & vbCrLf
#Else
    sql = sql & " ORDER BY " & cSQLOrderString & vbCrLf    '�C���A�G���[�A�x���A����̏�
    '//�ȍ~�̂n�q�c�d�q��
    Select Case cboSort.ListIndex
    Case eSort.eImportSeq
        sql = sql & ",FIINDT,FISEQN" & vbCrLf
    Case eSort.eKeiyakusha
        sql = sql & ",FIITKB,FIKYCD,FIKSCD,FIPGNO,FIFKDT,FIHGCD,FISEQN" & vbCrLf
    Case Else
    End Select
#End If
    sql = sql & ")" & vbCrLf
    pMakeSQLReadDataDetail = sql
End Function

Private Sub cmdErrList_Click()
    If True = pSpreadCheckAndUpdate(sprTotal.Tag = mSprTotal.RowEdit Or sprDetail.Tag = mSprDetail.RowEdit) Then
        Exit Sub
    End If
    Dim reg As New RegistryClass
    Dim sql As String
    Load rptFurikaeYoteiImport
    With rptFurikaeYoteiImport
        .lblSort.Caption = "�\�����F " & cboSort.Text
        .documentName = mCaption
        '//�����Őݒ�����Ă��ύX�ł��Ȃ��I
        '.PageSettings.PaperSize = vbPRPSA4
        '.PageSettings.Orientation = ddOPortrait
        .adoData.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=" & reg.DbPassword & _
                                    ";Persist Security Info=True;User ID=" & reg.DbUserName & _
                                                           ";Data Source=" & reg.DbDatabaseName
        sql = "SELECT * FROM (" & vbCrLf
        sql = sql & "SELECT " & vbCrLf
#If SHORT_MSG Then
        sql = sql & " DECODE(FIERROR,-3,'�捞',-2,'�C��',-1,'�ُ�',0,'����',1,'�x��','��O') as FIERRNM," & vbCrLf
#Else
        sql = sql & " CASE WHEN FIERROR = -2 THEN " & gdDBS.ColumnDataSet(cEditDataMsg, vEnd:=True) & vbCrLf
        sql = sql & "      WHEN FIERROR = -3 THEN " & gdDBS.ColumnDataSet(cImportMsg, vEnd:=True) & vbCrLf
'//2006/06/16 ���׍폜�Ή�
        sql = sql & "      WHEN FIERROR = -4 THEN " & gdDBS.ColumnDataSet(cDeleteMsg, vEnd:=True) & vbCrLf
        sql = sql & "      WHEN FIERROR IN(-1,+0,+1) THEN " & vbCrLf
        sql = sql & "           DECODE(FIERROR," & vbCrLf
        sql = sql & "               -1,'�ُ�'," & vbCrLf
        sql = sql & "               +0,'����'," & vbCrLf
        sql = sql & "               +1,'�x��'," & vbCrLf
        sql = sql & "               NULL" & vbCrLf
        sql = sql & "           ) || ' => ' || " & vbCrLf
        sql = sql & "       DECODE(FIOKFG," & vbCrLf
        sql = sql & "               " & mYimp.updInvalid & ",'" & mYimp.mUpdateMessage(mYimp.updInvalid) & "'," & vbCrLf
        sql = sql & "               " & mYimp.updWarnErr & ",'" & mYimp.mUpdateMessage(mYimp.updWarnErr) & "'," & vbCrLf
        sql = sql & "               " & mYimp.updNormal & ",'" & mYimp.mUpdateMessage(mYimp.updNormal) & "'," & vbCrLf
        sql = sql & "               " & mYimp.updWarnUpd & ",'" & mYimp.mUpdateMessage(mYimp.updWarnUpd) & "'," & vbCrLf
        '//����ȃf�[�^�͖���
        'sql = sql & "               " & mYimp.updResetCancel & ",'" & mYimp.mUpdateMessage(mYimp.updResetCancel) & "'," & vbCrLf
        sql = sql & "               '�������ʂ�����ł��܂���B'" & vbCrLf
        sql = sql & "           )" & vbCrLf
        sql = sql & "      ELSE                             '��O => �������ʂ�����ł��܂���B'" & vbCrLf
        sql = sql & " END as FIERRNM," & vbCrLf
#End If
        sql = sql & " FIRKBN,TO_CHAR(FIINDT,'yyyy/mm/dd hh24:mi:ss') FIINDT,FISEQN,"
        sql = sql & " TO_CHAR(TO_DATE(FIFKDT,'yyyymmdd'),'yyyy/mm/dd') fifkdt," & vbCrLf
        sql = sql & "(SELECT ABITCD " & vbCrLf
        sql = sql & " FROM taItakushaMaster b " & vbCrLf
        sql = sql & " WHERE a.FIITKB = b.ABITKB" & vbCrLf
        sql = sql & " ) ABITCD," & vbCrLf
        sql = sql & " FIKYCD,FIKSCD,FIPGNO,FIHGCD," & vbCrLf
        sql = sql & " DECODE(NVL(FIKYFG,0),0,FIHKKG,DECODE(NVL(FIHKKG,0),0,NULL,FIHKKG)) FIHKKG," & vbCrLf
        sql = sql & " DECODE(NVL(FIKYFG,0),0,NULL,'���') FIKYFG," & vbCrLf
        sql = sql & " FIHKCT,FIKYCT," & vbCrLf
        sql = sql & " FIITKB || FIKYCD || FIKSCD || FIPGNO || FIFKDT FIGROUP," & vbCrLf
        sql = sql & mYimp.StatusColumns("," & vbCrLf, Len("," & vbCrLf))
        sql = sql & " FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
        sql = sql & " WHERE FIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss') " & vbCrLf
        sql = sql & "   AND       (" & cInSQLString & ") IN(" & vbCrLf
        sql = sql & "       SELECT " & cInSQLString & vbCrLf
        sql = sql & "       FROM " & mYimp.TfFurikaeImport & " b " & vbCrLf
        sql = sql & "       WHERE a.FIINDT = b.FIINDT" & vbCrLf
        sql = sql & "         AND b.FIERROR <> " & mYimp.errNormal & vbCrLf
        sql = sql & "      )" & vbCrLf
        '//�ȍ~�̂n�q�c�d�q��
        sql = sql & " ORDER BY " & cSQLOrderString & vbCrLf    '�C���A�G���[�A�x���A����̏�
        Select Case cboSort.ListIndex
        Case eSort.eImportSeq
            sql = sql & " ,FIINDT,FISEQN" & vbCrLf
        Case eSort.eKeiyakusha
            'sql = sql & " ORDER BY FIINDT,FIITKB,FIKYCD,FIKSCD,FIPGNO,DECODE(FIRKBN,-1,999,FIRKBN),FISEQN" & vbCrLf
            sql = sql & " ,FIINDT,FIITKB,FIKYCD,FIKSCD,FIPGNO,FIFKDT,FIHGCD,FISEQN" & vbCrLf
        Case Else
        End Select
        sql = sql & ")"
        .adoData.Source = sql
        Call .adoData.Refresh
'        .mTotalCnt = .adoData.Recordset.RecordCount
        Call .Show
    End With
End Sub

Private Sub cmdImport_Click()
    If True = pSpreadCheckAndUpdate(sprTotal.Tag = mSprTotal.RowEdit Or sprDetail.Tag = mSprDetail.RowEdit) Then
        Exit Sub
    End If
    '//�{�^���̃R���g���[��
    If -1 <> pAbortButton(cmdImport, cBtnImport) Then
        Exit Sub
    End If
    cmdImport.Caption = cBtnCancel
    '//�R�}���h�E�{�^������
    Call pLockedControl(False, cmdImport)

    Dim file As New FileClass
    
    dlgFile.DialogTitle = "�t�@�C�����J��(" & mCaption & ")"
    dlgFile.FileName = mReg.InputFileName(mCaption)
    If IsEmpty(file.OpenDialog(dlgFile)) Then
        GoTo cmdImport_ClickAbort:
        Exit Sub
    End If
    '//�U���\��\�f�[�^���C���|�[�g
    Dim FurikaeDetail As tpFurikaeDetail
    Dim FurikaeTotal  As tpFurikaeTotal
    Dim fp As Integer
    Dim ms As New MouseClass
    Call ms.Start
    
    fp = FreeFile
    Open dlgFile.FileName For Random Access Read As #fp Len = Len(FurikaeDetail)
    fraProgressBar.Visible = True
    pgrProgressBar.Max = LOF(fp) / Len(FurikaeDetail)
    '//�t�@�C���T�C�Y���Ⴄ�ꍇ�̌x�����b�Z�[�W
    If pgrProgressBar.Max <> Int(pgrProgressBar.Max) Then
        If (LOF(fp) - 1) / Len(FurikaeDetail) <> Int((LOF(fp) - 1) / Len(FurikaeDetail)) Then
            '/�������s����Ƃc�a�����������Ȃ�̂Œ��~����
            Close #fp
            Call gdDBS.MsgBox("�w�肳�ꂽ�t�@�C��(" & dlgFile.FileName & ")���ُ�ł��B" & vbCrLf & vbCrLf & "�����𑱍s�o���܂���B", vbCritical + vbOKOnly, mCaption)
            GoTo cmdImport_ClickAbort
            Exit Sub
        End If
    End If

    On Error GoTo cmdImport_ClickError
        
    Call gdDBS.AutoLogOut(mCaption, "�捞�������J�n����܂����B")
    
#If ORA_DEBUG = 1 Then
    Dim sql As String, insDate As String, dyn As OraDynaset
#Else
    Dim sql As String, insDate As String, dyn As Object
#End If
    Dim updCnt As Long, insCnt As Long, recCnt As Long
    
'//2006/06/16 �_��Ҕԍ������̃p���`�f�[�^�Ή�
    Dim BadNo As Long
    BadNo = cFIKYCD_BadStart
    
    insDate = gdDBS.sysDate()
    
    Call gdDBS.Database.BeginTrans
    '///////////////////////////////////////////////
    '//�V�[�P���X���P�Ԃ���Ƀ��Z�b�g
    sql = "declare begin ResetSequence('sqImportSeq',1); end;"
    Call gdDBS.Database.ExecuteSQL(sql)
    
    Do While Loc(fp) < LOF(fp) / Len(FurikaeDetail)
        DoEvents
        If mAbort Then
            GoTo cmdImport_ClickError
        End If
        Get #fp, , FurikaeDetail
        recCnt = Loc(fp)
'//2006/06/16 �_��Ҕԍ������̃p���`�f�[�^�Ή�
        If "" = Trim(FurikaeDetail.KeiyakuNo) Then
            FurikaeDetail.KeiyakuNo = BadNo
        End If
        If "" = Trim(FurikaeDetail.KyoshitsuNo) Then
            FurikaeDetail.KyoshitsuNo = "000"       '//���͂��Ȃ��Ȃ��悤�ɂ킴�� "000"
        End If
        If "" = Trim(FurikaeDetail.PageNumber) Then
            FurikaeDetail.PageNumber = "00"         '//���͂��Ȃ��Ȃ��悤�ɂ킴�� "00"
        End If
'//2006/05/17 �U�֗\����̖����f�[�^�����邽�߃T�[�o�[�̖{����ݒ肷��
        If "" = Trim(FurikaeDetail.FurikaeDate) Then
            FurikaeDetail.FurikaeDate = Format(insDate, "yyyymmdd")
        End If
        If FurikaeDetail.CancelFlag = mYimp.TotalTextKubun Then
'//2006/06/16 �_��Ҕԍ������̃p���`�f�[�^�Ή��F�g�[�^�����R�[�h���ɔԍ����Z
            BadNo = BadNo + 1
            '//�g�[�^�����R�[�h�Ȃ̂ŃR�s�[
            LSet FurikaeTotal = FurikaeDetail
        End If
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = _
            "�c��" & Right(String(7, " ") & Format(pgrProgressBar.Max - Loc(fp), "#,##0"), 7) & " ��"
        pgrProgressBar.Value = IIf(Loc(fp) <= pgrProgressBar.Max, Loc(fp), pgrProgressBar.Max)
        sql = "SELECT FIMCDT"
        sql = sql & " FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
        sql = sql & " WHERE FIINDT = TO_DATE(" & gdDBS.ColumnDataSet(insDate, vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss')" & vbCrLf
        sql = sql & "   AND FISEQN > 0" & vbCrLf
        sql = sql & "   AND FIKYCD = " & gdDBS.ColumnDataSet(FurikaeDetail.KeiyakuNo, vEnd:=True) & vbCrLf
        sql = sql & "   AND FIKSCD = " & gdDBS.ColumnDataSet(FurikaeDetail.KyoshitsuNo, vEnd:=True) & vbCrLf
        sql = sql & "   AND FIPGNO = " & gdDBS.ColumnDataSet(FurikaeDetail.PageNumber, "I", vEnd:=True) & vbCrLf
        sql = sql & "   AND FIFKDT = " & gdDBS.ColumnDataSet(FurikaeDetail.FurikaeDate, "L", vEnd:=True) & vbCrLf
        If FurikaeDetail.CancelFlag = mYimp.TotalTextKubun Then
            sql = sql & " AND FIRKBN = " & gdDBS.ColumnDataSet(mYimp.RecordIsTotal, "I", vEnd:=True) & vbCrLf    '//���R�[�h�敪
        Else
            sql = sql & " AND FIRKBN <> " & gdDBS.ColumnDataSet(mYimp.RecordIsTotal, "I", vEnd:=True) & vbCrLf      '//���R�[�h�敪
            sql = sql & " AND FIHGCD = " & gdDBS.ColumnDataSet(FurikaeDetail.HogoshaNo, vEnd:=True) & vbCrLf
        End If
#If ORA_DEBUG = 1 Then
        Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
        Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
        If Not dyn.EOF Then
'//2006/04/27 �f�[�^�����𐳊m�ɂ��邽�߂ɂƂɂ����f�[�^������΍X�V����
            '//�X�V�����݂�F����e�L�X�g���ɓ����f�[�^������H
            sql = "UPDATE " & mYimp.TfFurikaeImport & " SET " & vbCrLf
            '//������������Ȃ�X�V�FDetail �A Total �ǂ���ł��n�j
            If Val(dyn.Fields("FIMCDT")) < Val(FurikaeDetail.MochikomiBi) Then
        
                If FurikaeDetail.CancelFlag = mYimp.TotalTextKubun Then
                    sql = sql & "FIHKCT = " & gdDBS.ColumnDataSet(gdDBS.Nz(FurikaeTotal.DetailCnt, 0), "I") & vbCrLf
                    sql = sql & "FIKYCT = " & gdDBS.ColumnDataSet(gdDBS.Nz(FurikaeTotal.CancelCnt, 0), "I") & vbCrLf
                    sql = sql & "FIHKKG = " & gdDBS.ColumnDataSet(gdDBS.Nz(FurikaeTotal.DetailGaku, 0), "L") & vbCrLf
                Else
'//2006/04/27 �������`�l�����ڒǉ��F�S�p�X�y�[�X�𔼊p�X�y�[�X�ɕϊ�
                    sql = sql & "FIKZNM = " & gdDBS.ColumnDataSet(Replace(FurikaeDetail.KouzaName, "�@", " ")) & vbCrLf
                    sql = sql & "FIHKKG = " & gdDBS.ColumnDataSet(gdDBS.Nz(FurikaeDetail.HenkouGaku, 0), "L") & vbCrLf
                    '//�L�����Z���t���O BLANK ����
                    sql = sql & "FIKYFG = " & gdDBS.ColumnDataSet(gdDBS.Nz(FurikaeDetail.CancelFlag, 0), "I") & vbCrLf
                End If
                sql = sql & "FIERROR = " & gdDBS.ColumnDataSet(mYimp.errImport) & vbCrLf
                sql = sql & "FIMCDT = " & gdDBS.ColumnDataSet(FurikaeDetail.MochikomiBi, "L") & vbCrLf
            '//���������O���Ȃ�X�V�����̂ݍX�V�FDetail �A Total �ǂ���ł��n�j
            End If
'//2006/04/27 �P�t�@�C�����̎捞�񐔁F�������݂���
            sql = sql & "FIICNT = FIICNT + 1," & vbCrLf
            sql = sql & "FIUPDT = SYSDATE" & vbCrLf
            sql = sql & " WHERE FIINDT = TO_DATE(" & gdDBS.ColumnDataSet(insDate, vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss')" & vbCrLf
            sql = sql & "   AND FISEQN > 0" & vbCrLf
            sql = sql & "   AND FIKYCD = " & gdDBS.ColumnDataSet(FurikaeDetail.KeiyakuNo, vEnd:=True) & vbCrLf
            sql = sql & "   AND FIKSCD = " & gdDBS.ColumnDataSet(FurikaeDetail.KyoshitsuNo, vEnd:=True) & vbCrLf
            sql = sql & "   AND FIPGNO = " & gdDBS.ColumnDataSet(FurikaeDetail.PageNumber, "I", vEnd:=True) & vbCrLf
            sql = sql & "   AND FIFKDT = " & gdDBS.ColumnDataSet(FurikaeDetail.FurikaeDate, "L", vEnd:=True) & vbCrLf
            If FurikaeDetail.CancelFlag = mYimp.TotalTextKubun Then
                sql = sql & " AND FIRKBN = " & gdDBS.ColumnDataSet(mYimp.RecordIsTotal, "I", vEnd:=True) & vbCrLf    '//���R�[�h�敪
            Else
                sql = sql & " AND FIRKBN <> " & gdDBS.ColumnDataSet(mYimp.RecordIsTotal, "I", vEnd:=True) & vbCrLf      '//���R�[�h�敪
                sql = sql & " AND FIHGCD = " & gdDBS.ColumnDataSet(FurikaeDetail.HogoshaNo, vEnd:=True) & vbCrLf
            End If
            Call gdDBS.Database.ExecuteSQL(sql)
            updCnt = updCnt + 1&
        Else
            insCnt = insCnt + 1&
            '//�X�V�ł��Ȃ������̂ő}�������݂�
            '//�f�[�^���e�[�u���ɑ}��
            sql = "INSERT INTO " & mYimp.TfFurikaeImport & "(" & vbCrLf
            sql = sql & "FIINDT,"   '//A=  �捞��
            sql = sql & "FISEQN,"   '//A=  �捞SEQNO
            sql = sql & "FIITKB,"   '//A=  �ϑ��ҋ敪
            sql = sql & "FIKYCD,"   '//A=  �_��Ҕԍ�
            sql = sql & "FIKSCD,"   '//A=  �����ԍ�
            sql = sql & "FIPGNO,"   '//A=  �y�[�W�ԍ�
            sql = sql & "FIFKDT,"   '//A=  �U�֓�
            sql = sql & "FIRKBN,"   '//B/T=���R�[�h�敪 �O�����ׁA�P�����v
            sql = sql & "FIHKCT,"   '//  T=�ύX����
            sql = sql & "FIKYCT,"   '//  T=��񌏐�
            sql = sql & "FIHGCD,"   '//B=  �ی�Ҕԍ�
'//2006/04/27 �������`�l�����ڒǉ�
            sql = sql & "FIKZNM,"   '//  T=�ی�ҁE�������`�l��
            sql = sql & "FIHKKG,"   '//B/T=�ύX����z
            sql = sql & "FIKYFG,"   '//B=  ���t���O
            sql = sql & "FIERROR,"
            sql = sql & "FIMCDT,"   '//������
'//2006/04/27 �P�t�@�C�����̎捞�񐔁F�������݂���
            sql = sql & "FIICNT," & vbCrLf
            sql = sql & "FIUSID,"   '//A=  �X�V��
            sql = sql & "FIUPDT,"   '//A=  �X�V��
            sql = sql & "FIOKFG " & vbCrLf  '//�捞�n�j�t���O
            sql = sql & ")VALUES(" & vbCrLf
            sql = sql & "TO_DATE(" & gdDBS.ColumnDataSet(insDate, vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss')," & vbCrLf
            sql = sql & "sqImportSeq.NEXTVAL," & vbCrLf
'//2006/06/16 �_��Ҕԍ������̃p���`�f�[�^�Ή�
'//            sql = sql & "(SELECT ABITKB FROM taItakushaMaster WHERE ABKYTP = '" & Left(FurikaeDetail.KeiyakuNo, 1) & "')," & vbCrLf
            sql = sql & "(SELECT DECODE(MAX(ABITKB),NULL,'" & cFIITKB_BadCode & "',MAX(ABITKB)) FROM taItakushaMaster WHERE ABKYTP = '" & Left(FurikaeDetail.KeiyakuNo, 1) & "')," & vbCrLf
            sql = sql & gdDBS.ColumnDataSet(FurikaeDetail.KeiyakuNo) & vbCrLf
            sql = sql & gdDBS.ColumnDataSet(FurikaeDetail.KyoshitsuNo) & vbCrLf
            sql = sql & gdDBS.ColumnDataSet(FurikaeDetail.PageNumber, "I") & vbCrLf
            sql = sql & gdDBS.ColumnDataSet(FurikaeDetail.FurikaeDate, "L") & vbCrLf
            If FurikaeDetail.CancelFlag = mYimp.TotalTextKubun Then
                sql = sql & gdDBS.ColumnDataSet(mYimp.RecordIsTotal, "I") & vbCrLf    '//���R�[�h�敪
                sql = sql & gdDBS.ColumnDataSet(gdDBS.Nz(FurikaeTotal.DetailCnt, 0), "I") & vbCrLf
                sql = sql & gdDBS.ColumnDataSet(gdDBS.Nz(FurikaeTotal.CancelCnt, 0), "I") & vbCrLf
                sql = sql & "NULL," & vbCrLf
'//2006/04/27 �������`�l�����ڒǉ�
                sql = sql & "NULL," & vbCrLf
                sql = sql & gdDBS.ColumnDataSet(gdDBS.Nz(FurikaeTotal.DetailGaku, 0), "L") & vbCrLf
                sql = sql & "NULL," & vbCrLf
            Else
                sql = sql & "0," & vbCrLf     '//���R�[�h�敪�F��Őe�̂r�d�p��������.
                sql = sql & "NULL," & vbCrLf
                sql = sql & "NULL," & vbCrLf
                sql = sql & gdDBS.ColumnDataSet(FurikaeDetail.HogoshaNo) & vbCrLf
'//2006/04/27 �������`�l�����ڒǉ��F�S�p�X�y�[�X�𔼊p�X�y�[�X�ɕϊ�
                sql = sql & gdDBS.ColumnDataSet(Replace(FurikaeDetail.KouzaName, "�@", " ")) & vbCrLf
                sql = sql & gdDBS.ColumnDataSet(gdDBS.Nz(FurikaeDetail.HenkouGaku, 0), "L") & vbCrLf
                '//�L�����Z���t���O BLANK ����
                sql = sql & gdDBS.ColumnDataSet(gdDBS.Nz(FurikaeDetail.CancelFlag, 0), "I") & vbCrLf
            End If
            sql = sql & gdDBS.ColumnDataSet(mYimp.errImport) & vbCrLf
            sql = sql & gdDBS.ColumnDataSet(FurikaeDetail.MochikomiBi, "L") & vbCrLf
'//�P�t�@�C�����̎捞�񐔁F�������݂���
            sql = sql & " 1," & vbCrLf
            sql = sql & gdDBS.ColumnDataSet(gdDBS.LoginUserName)
            sql = sql & "SYSDATE,"
            sql = sql & gdDBS.ColumnDataSet(mYimp.updNormal, "I", vEnd:=True)
            sql = sql & ")"
            Call gdDBS.Database.ExecuteSQL(sql)
        End If
        Call dyn.Close
        Set dyn = Nothing
    Loop
    Close #fp
    sql = "UPDATE " & mYimp.TfFurikaeImport & " a SET " & vbCrLf
    sql = sql & " FIRKBN = (" & vbCrLf
    sql = sql & "     SELECT FISEQN FROM " & mYimp.TfFurikaeImport & " b " & vbCrLf
    sql = sql & "     WHERE b.FIINDT = a.FIINDT" & vbCrLf
    sql = sql & "       AND b.FIKYCD = a.FIKYCD" & vbCrLf
    sql = sql & "       AND b.FIKSCD = a.FIKSCD" & vbCrLf
    sql = sql & "       AND b.FIPGNO = a.FIPGNO" & vbCrLf
    sql = sql & "       AND b.FIFKDT = a.FIFKDT" & vbCrLf
    sql = sql & "       AND b.FIRKBN = " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & "     )" & vbCrLf
    sql = sql & " WHERE a.FIINDT = TO_DATE(" & gdDBS.ColumnDataSet(insDate, vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss')" & vbCrLf
    sql = sql & "   AND a.FIRKBN <> " & mYimp.RecordIsTotal & vbCrLf
    Call gdDBS.Database.ExecuteSQL(sql)
     '//�X�e�[�^�X�s�̐���E����
    stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�捞����(" & recCnt & "��)"
    pgrProgressBar.Value = pgrProgressBar.Max
   '//�U���\��\�f�[�^�̈ʒu�����W�X�g���ɕۊ�
    mReg.InputFileName(mCaption) = dlgFile.FileName
    '//�捞�f�[�^�̃o�b�N�A�b�v
    Call gBackupTextData(dlgFile.FileName)

    Call gdDBS.Database.CommitTrans
    
    Call gdDBS.AutoLogOut(mCaption, "�捞����=[" & insDate & "]�� " & recCnt & " ���i�ǉ�=" & insCnt & " / �d��=" & updCnt & "�j�̃f�[�^����荞�܂�܂����B")
    
    '//�捞���ʂ��R���{�{�b�N�X�ɃZ�b�g
    Call pMakeComboBox

cmdImport_ClickAbort:
    '//���ׂĂ̒�`�����Z�b�g
    Set file = Nothing
    Set ms = Nothing
    cmdImport.Caption = cBtnImport
    fraProgressBar.Visible = False
    Call pLockedControl(True)
    Exit Sub
cmdImport_ClickError:
    '//�X�e�[�^�X�s�̐���E����
    'cmdImport.Caption = cBtnImport
    Call gdDBS.Database.Rollback
    Call gdDBS.ErrorCheck       '//�G���[�g���b�v
    If Err Then
        Dim errCode As Integer, errMsg As String
        If gdDBS.Database.LastServerErr Then
            errCode = gdDBS.Database.LastServerErr
            errMsg = gdDBS.Database.LastServerErrText
        Else
            errCode = Err
            errMsg = Error
        End If
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�捞�G���[(" & errCode & ")"
        Call gdDBS.AutoLogOut(mCaption, recCnt & "���ڂŃG���[�������������ߎ捞�����͒��~����܂����B(Error=" & errMsg & ")")
    Else
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�捞���f"
        Call gdDBS.AutoLogOut(mCaption, "�捞�����͒��~����܂����B")
    End If
    'Call pLockedControl(True)
    GoTo cmdImport_ClickAbort:
End Sub

Private Sub pMakeComboBox()
    Dim ms As New MouseClass
    Call ms.Start
    '//�R�}���h�E�{�^������
    Call pLockedControl(False)
'    Dim sql As String, dyn As OraDynaset, MaxDay As Variant
    Dim sql As String, dyn As Object, MaxDay As Variant
    sql = "SELECT DISTINCT TO_CHAR(FIINDT,'yyyy/mm/dd hh24:mi:ss') FIINDT_A"
    sql = sql & " FROM " & mYimp.TfFurikaeImport
    sql = sql & " ORDER BY FIINDT_A"
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
    Call cboImpDate.Clear
    Do Until dyn.EOF()
        Call cboImpDate.AddItem(dyn.Fields("FIINDT_A"))
        'cboImpDate.ItemData(cboImpDate.NewIndex) = dyn.Fields("CIINDT_B")
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    If cboImpDate.ListCount Then
        cboImpDate.ListIndex = cboImpDate.ListCount - 1
    Else
        sprTotal.MaxRows = 0
    End If
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
End Sub
        
Private Sub pUpdateDetail()
'//���ӁI
'// mSprDetail/mSprTotal & eSprDetail/eSprTotal �̍\�������Ă���̂Œ���
    Dim sql As String
    Dim Row As Long
    For Row = 1 To sprDetail.MaxRows
        If Val(mSprDetail.Value(eSprDetail.eEditFlag, Row)) = mSprDetail.RowEdit Then
            sql = "UPDATE " & mYimp.TfFurikaeImport & " SET " & vbCrLf
            sql = sql & "FIERROR = " & gdDBS.ColumnDataSet(mSprDetail.Value(eSprDetail.eErrorFlag, Row), "I") & vbCrLf
            sql = sql & "FIHGCD  = " & gdDBS.ColumnDataSet(mSprDetail.Value(eSprDetail.eHogoshaNo, Row)) & vbCrLf
'//2006/04/27 �������`�l�����ڒǉ�
            sql = sql & "FIKZNM  = " & gdDBS.ColumnDataSet(mSprDetail.Value(eSprDetail.eImportKouza, Row)) & vbCrLf
'//2006/04/26 ���z�Ȃ̂� NULL �ł͖��� �u�O�v��������
            sql = sql & "FIHKKG  = " & gdDBS.ColumnDataSet(gdDBS.Nz(mSprDetail.Value(eSprDetail.eHenkoGaku, Row), 0), "L") & vbCrLf
            sql = sql & "FIKYFG  = " & gdDBS.ColumnDataSet(mSprDetail.Value(eSprDetail.eCancelFlag, Row), "I") & vbCrLf
            sql = sql & "FIUSID  = " & gdDBS.ColumnDataSet(gdDBS.LoginUserName) & vbCrLf
            sql = sql & "FIUPDT  = SYSDATE" & vbCrLf
            sql = sql & " WHERE FIINDT = TO_DATE(" & gdDBS.ColumnDataSet(mSprDetail.Value(eSprDetail.eImpDate, Row), vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss') " & vbCrLf
            sql = sql & "   AND FISEQN = " & gdDBS.ColumnDataSet(mSprDetail.Value(eSprDetail.eImpSEQ, Row), "L", vEnd:=True) & vbCrLf
            Call gdDBS.Database.ExecuteSQL(sql)
            '//�C���t���O���Z�b�g
            mSprDetail.Value(eSprDetail.eEditFlag, Row) = mSprDetail.RowNonEdit
        End If
    Next Row
    sprDetail.Tag = mSprDetail.RowNonEdit
End Sub

Private Sub pUpdateTotal()
    Dim sql As String, updCnt As Long
    Dim Row As Long
    
    For Row = 1 To sprTotal.MaxRows
        If Val(mSprTotal.Value(eSprTotal.eEditFlag, Row)) = mSprTotal.RowEdit _
        Or Val(mSprTotal.Value(eSprTotal.eEditFlag, Row)) = mSprTotal.RowEditHeader Then
            '//���v�s�̍X�V
            sql = "UPDATE " & mYimp.TfFurikaeImport & " SET " & vbCrLf
            sql = sql & "FIERROR = " & gdDBS.ColumnDataSet(mSprTotal.Value(eSprTotal.eErrorFlag, Row), "I") & vbCrLf
'//2006/04/26 �����A���z�Ȃ̂� NULL �ł͖��� �u�O�v��������
            sql = sql & "FIHKCT  = " & gdDBS.ColumnDataSet(gdDBS.Nz(mSprTotal.Value(eSprTotal.eHenkoCount, Row), 0), "I") & vbCrLf
            sql = sql & "FIHKKG  = " & gdDBS.ColumnDataSet(gdDBS.Nz(mSprTotal.Value(eSprTotal.eHenkoKingaku, Row), 0), "L") & vbCrLf
            sql = sql & "FIKYCT  = " & gdDBS.ColumnDataSet(gdDBS.Nz(mSprTotal.Value(eSprTotal.eCancelCount, Row), 0), "I") & vbCrLf
            sql = sql & "FIUSID  = " & gdDBS.ColumnDataSet(gdDBS.LoginUserName) & vbCrLf
            sql = sql & "FIUPDT  = SYSDATE" & vbCrLf
            sql = sql & " WHERE FIINDT = TO_DATE(" & gdDBS.ColumnDataSet(mSprTotal.Value(eSprTotal.eImpDate, Row), vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss') " & vbCrLf
            sql = sql & "   AND FISEQN = " & gdDBS.ColumnDataSet(mSprTotal.Value(eSprTotal.eImpSEQ, Row), "L", vEnd:=True) & vbCrLf
            updCnt = gdDBS.Database.ExecuteSQL(sql)
            '//���v�s�Ɋ֘A�������׍s�̍X�V
            If Val(mSprTotal.Value(eSprTotal.eEditFlag, Row)) = mSprTotal.RowEditHeader Then
                sql = "UPDATE " & mYimp.TfFurikaeImport & " SET " & vbCrLf
                sql = sql & "FIERROR = " & gdDBS.ColumnDataSet(mSprTotal.Value(eSprTotal.eErrorFlag, Row), "I") & vbCrLf
                sql = sql & "FIITKB  = " & gdDBS.ColumnDataSet(mSprTotal.Value(eSprTotal.eItakuCode, Row)) & vbCrLf
                sql = sql & "FIKYCD  = " & gdDBS.ColumnDataSet(mSprTotal.Value(eSprTotal.eKeiyakuCode, Row)) & vbCrLf
                sql = sql & "FIKSCD  = " & gdDBS.ColumnDataSet(mSprTotal.Value(eSprTotal.eKyoshitsuNo, Row)) & vbCrLf
                sql = sql & "FIPGNO  = " & gdDBS.ColumnDataSet(mSprTotal.Value(eSprTotal.ePageNumber, Row), "I") & vbCrLf
                '//�uyyyy/mm/dd�v�œ��͂��Ă���̂� yyyymmdd �ɕϊ������t�̐������`�F�b�N
                sql = sql & "FIFKDT  = TO_CHAR(TO_DATE(" & gdDBS.ColumnDataSet(mSprTotal.Value(eSprTotal.eFirukaeDate, Row), vEnd:=True) & ",'yyyy/mm/dd'),'yyyymmdd')," & vbCrLf
                sql = sql & "FIUSID  = " & gdDBS.ColumnDataSet(gdDBS.LoginUserName) & vbCrLf
                sql = sql & "FIUPDT  = SYSDATE" & vbCrLf
'z 2006/06/19 ���׍s�X�V�Ƀo�O������̂ŕύX�F FIRKBN ���Q�Ƃ���.
'z                sql = sql & " WHERE (" & cInSQLString & ") IN (" & vbCrLf
'z                sql = sql & "   SELECT " & cInSQLString & vbCrLf
'z                sql = sql & "   FROM " & mYimp.TfFurikaeImport & vbCrLf
'z                sql = sql & "   WHERE FIINDT = TO_DATE(" & gdDBS.ColumnDataSet(mSprTotal.Value(eSprTotal.eImpDate, Row), vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss') " & vbCrLf
'z                sql = sql & "     AND FISEQN = " & gdDBS.ColumnDataSet(mSprTotal.Value(eSprTotal.eImpSEQ, Row), "L", vEnd:=True) & vbCrLf
'z                sql = sql & "  )" & vbCrLf
                sql = sql & " WHERE FIINDT = TO_DATE(" & gdDBS.ColumnDataSet(mSprTotal.Value(eSprTotal.eImpDate, Row), vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss') " & vbCrLf
                sql = sql & "   AND(FISEQN = " & gdDBS.ColumnDataSet(mSprTotal.Value(eSprTotal.eImpSEQ, Row), "L", vEnd:=True) & vbCrLf
                sql = sql & "    OR FIRKBN = " & gdDBS.ColumnDataSet(mSprTotal.Value(eSprTotal.eImpSEQ, Row), "L", vEnd:=True) & vbCrLf
                sql = sql & "   )" & vbCrLf
                updCnt = gdDBS.Database.ExecuteSQL(sql)
            End If
            '//�C���t���O���Z�b�g
            mSprTotal.Value(eSprTotal.eEditFlag, Row) = mSprTotal.RowNonEdit
        End If
    Next Row
    sprTotal.Tag = mSprTotal.RowNonEdit
End Sub

Private Sub cmdSprUpdate_Click()
    If -1 <> pAbortButton(cmdSprUpdate, cBtnSprUpdate) Then
        Exit Sub
    End If
    cmdSprUpdate.Caption = cBtnCancel
    '//�R�}���h�E�{�^������
    Call pLockedControl(False, cmdSprUpdate)
    Dim ms As New MouseClass
    Call ms.Start
    
    Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] �̍X�V���J�n����܂����B")
    
'    On Error GoTo cmdSprUpdate_ClickError:
    Call gdDBS.Database.BeginTrans
    
    '//���׃��R�[�h�̍X�V
    If sprDetail.Tag = mSprDetail.RowEdit Then
        Call pUpdateDetail
    End If
    '//���v���R�[�h�̍X�V
    If sprTotal.Tag = mSprTotal.RowEdit Then
        Call pUpdateTotal
    End If
    Call gdDBS.Database.CommitTrans
    '//�X�e�[�^�X�s�̐���E����
    stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�X�V����"
    Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "]  �̍X�V���������܂����B")
    
    '//�{�^����߂�
    cmdSprUpdate.Caption = cBtnSprUpdate
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
    Exit Sub
cmdSprUpdate_ClickError:
    Call gdDBS.Database.Rollback
    If Err Then
        Dim errCode As Integer, errMsg As String
        If gdDBS.Database.LastServerErr Then
            errCode = gdDBS.Database.LastServerErr
            errMsg = gdDBS.Database.LastServerErrText
        Else
            errCode = Err
            errMsg = Error
        End If
        fraProgressBar.Visible = False
        '//�X�e�[�^�X�s�̐���E����
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�X�V�G���[(" & errCode & ")"
        Call gdDBS.AutoLogOut(mCaption, "�G���[�������������ߍX�V�͒��~����܂����B(Error=" & errMsg & ")")
        Call MsgBox("�X�V�Ώ� = [" & cboImpDate.Text & "]" & vbCrLf & _
                    "�̓G���[�������������ߍX�V�͒��~����܂����B" & vbCrLf & errMsg, _
                vbOKOnly + vbCritical, mCaption)
    Else
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�X�V���f"
        Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] �̍X�V�͒��~����܂����B")
    End If
    '//�{�^����߂�
    cmdUpdate.Caption = cBtnSprUpdate
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
End Sub

Private Sub cmdUpdate_Click()
    If True = pSpreadCheckAndUpdate(sprTotal.Tag = mSprTotal.RowEdit Or sprDetail.Tag = mSprDetail.RowEdit) Then
        Exit Sub
    End If
    If -1 <> pAbortButton(cmdUpdate, cBtnUpdate) Then
        Exit Sub
    End If
    cmdUpdate.Caption = cBtnCancel
    '//�R�}���h�E�{�^������
    Call pLockedControl(False, cmdUpdate)
    Dim ms As New MouseClass
    Call ms.Start
    
    Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] �̃}�X�^���f���J�n����܂����B")
    
    On Error GoTo cmdUpdate_ClickError:
    Call gdDBS.Database.BeginTrans
        
    '//�����̐ݒ�
    Dim CanDate As String
    CanDate = gdDBS.SystemUpdate("AANXKZ")
    CanDate = Format(DateSerial(Val(Mid(CanDate, 1, 4)), Val(Mid(CanDate, 5, 2)), Val(Mid(CanDate, 7, 2)) - 1), "yyyymmdd")
    '//�V�X�e���}�X�^�̎���U�֓�(�ی�҈�)�̑O��
    
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset, recCnt As Long, updCnt As Long
#Else
    Dim sql As String, dyn As Object, recCnt As Long, updCnt As Long
#End If
    '//////////////////////////////////////////////////////////
    '//�����Ŏg�p���鋤�ʂ� WHERE ����
    Dim Condition As String
    Condition = Condition & " AND FIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss') " & vbCrLf
    '// ----------------------------------------------------------------->>>��������
    Condition = Condition & " AND       (" & cInSQLString & ") NOT IN(" & vbCrLf
    Condition = Condition & "     SELECT " & cInSQLString & vbCrLf
    Condition = Condition & "     FROM " & mYimp.TfFurikaeImport & " b " & vbCrLf
    Condition = Condition & "     WHERE a.FIINDT = b.FIINDT" & vbCrLf
    Condition = Condition & "       AND b.FIOKFG <> " & mYimp.updNormal & vbCrLf      '//����łȂ��FNOT IN
    Condition = Condition & "    )" & vbCrLf
'//2006/06/16 ���׍폜�Ή�
    Condition = Condition & " AND FIERROR >= " & mYimp.errNormal & vbCrLf
    
    
    '//�P�O���[�v���ł��ׂĐ���(FIOKFG=0)�Ŗ����ƍX�V�͂��Ȃ�
    sql = "SELECT a.*" & vbCrLf
    sql = sql & " FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
    sql = sql & " WHERE 1 = 1" & vbCrLf '//���܂��Ȃ�
    sql = sql & Condition
    sql = sql & "  AND FIRKBN <> " & mYimp.RecordIsTotal & vbCrLf           '//���v���R�[�h�͕s�v
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    recCnt = dyn.RecordCount
    If dyn.EOF Then
        Call dyn.Close
        Set dyn = Nothing
        Call MsgBox("�捞���� [ " & cboImpDate.Text & " ]" & vbCrLf & "�Ƀ}�X�^���f���ׂ��f�[�^�͂���܂���B", vbOKOnly + vbInformation, mCaption)
        '//�{�^����߂�
        cmdUpdate.Caption = cBtnUpdate
        '//�R�}���h�E�{�^������
        Call pLockedControl(True)
        Exit Sub
    End If
    fraProgressBar.Visible = True
    pgrProgressBar.Max = recCnt
    Do Until dyn.EOF
        DoEvents
        If mAbort Then
            GoTo cmdUpdate_ClickError
        End If
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = _
            "�c��" & Right(String(7, " ") & Format(recCnt - dyn.RowPosition, "#,##0"), 7) & " ��"
        pgrProgressBar.Value = dyn.RowPosition
        '///////////////////////////////////////////////
        '//�U�֗\��f�[�^�ւ̍X�V
        '///////////////////////////////////////////////
        sql = "UPDATE tfFurikaeYoteiData SET " & vbCrLf
'frmKouzaFurikaeExport(�����U�փf�[�^�쐬) �ł̃f�[�^�� FASKGK ���o�͂��Ă���I
        sql = sql & " FASKGK = " & gdDBS.ColumnDataSet(dyn.Fields("FIHKKG"), "L") & vbCrLf
        'sql = sql & " FAHKGK = " & gdDBS.ColumnDataSet(dyn.Fields("FIHKGK"), "L") & vbCrLf
        sql = sql & " FAKYFG = " & gdDBS.ColumnDataSet(dyn.Fields("FIKYFG"), "I") & vbCrLf
        '//2003/02/03 �X�V��ԃt���O�ǉ�:0=DB�쐬,1=�\��쐬,2=�\��捞,3=�����쐬
        sql = sql & " FAUPFG = " & gdDBS.ColumnDataSet(eKouFuriKubun.YoteiImport, "I") & vbCrLf
        sql = sql & " FAUSID = " & gdDBS.ColumnDataSet(gcImportUserName) & vbCrLf  '//�X�V�҂h�c
        sql = sql & " FAUPDT = SYSDATE" & vbCrLf
        sql = sql & " WHERE FAITKB = " & gdDBS.ColumnDataSet(dyn.Fields("FIITKB"), vEnd:=True) & vbCrLf
        sql = sql & "   AND FAKYCD = " & gdDBS.ColumnDataSet(dyn.Fields("FIKYCD"), vEnd:=True) & vbCrLf
        sql = sql & "   AND FAKSCD = " & gdDBS.ColumnDataSet(dyn.Fields("FIKSCD"), vEnd:=True) & vbCrLf
        sql = sql & "   AND FAHGCD = " & gdDBS.ColumnDataSet(dyn.Fields("FIHGCD"), vEnd:=True) & vbCrLf
        sql = sql & "   AND FASQNO = " & gdDBS.ColumnDataSet(dyn.Fields("FIFKDT"), vEnd:=True) & vbCrLf
        updCnt = gdDBS.Database.ExecuteSQL(sql)
        '///////////////////////////////////////////////
        '//2003/02/03 �ی�҃}�X�^�ւ̍X�V
        '///////////////////////////////////////////////
        sql = "UPDATE tcHogoshaMaster SET " & vbCrLf
        sql = sql & " CASKGK = " & gdDBS.ColumnDataSet(dyn.Fields("FIHKKG"), "L") & vbCrLf
        sql = sql & " CAKYFG = " & gdDBS.ColumnDataSet(dyn.Fields("FIKYFG"), "I") & vbCrLf
        '//��񂳂ꂽ�̂Ō����U�֏I�����������̓��t�Ŗ��ߍ���
        If 0 <> Val(gdDBS.Nz(dyn.Fields("FIKYFG"))) Then
            'sql = sql & " CAFKED = TO_CHAR(SYSDATE,'YYYYMMDD')," & vbCrLf
            '//�擪�Őݒ肵�Ă���F�V�X�e���}�X�^�̎���U�֓�(�ی�҈�)�̑O��
            sql = sql & " CAFKED = " & gdDBS.ColumnDataSet(CanDate, "I") & vbCrLf
        End If
        sql = sql & " CAUSID = " & gdDBS.ColumnDataSet(gcImportUserName) & vbCrLf  '//�X�V�҂h�c
        sql = sql & " CAUPDT = SYSDATE" & vbCrLf
        sql = sql & " WHERE CAITKB = " & gdDBS.ColumnDataSet(dyn.Fields("FIITKB"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CAKYCD = " & gdDBS.ColumnDataSet(dyn.Fields("FIKYCD"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CAKSCD = " & gdDBS.ColumnDataSet(dyn.Fields("FIKSCD"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CAHGCD = " & gdDBS.ColumnDataSet(dyn.Fields("FIHGCD"), vEnd:=True) & vbCrLf
        sql = sql & "   AND " & gdDBS.ColumnDataSet(dyn.Fields("FIFKDT"), vEnd:=True) & _
                            " BETWEEN CAFKST AND CAFKED " & vbCrLf
        updCnt = gdDBS.Database.ExecuteSQL(sql)
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    Set dyn = Nothing
'//�}�X�^�[���f�̌����ڍׂ��擾����F�p���`�f�[�^�Ƃ̌����`�F�b�N�p
    Dim total(0 To 2) As Long
    Dim Detail(0 To 2) As Long
    Dim BadCnt(0 To 2) As Long
    '//���v�s���
    sql = "SELECT " & vbCrLf
    sql = sql & " SUM(DECODE(NVL(FIOKFG, 0)," & mYimp.updNormal & ",  NVL(FIICNT,0),0)) OK_CNT," & vbCrLf
    sql = sql & " SUM(DECODE(NVL(FIOKFG, 0)," & mYimp.updNormal & ",0,NVL(FIICNT,0)  )) NG_CNT," & vbCrLf
    sql = sql & " SUM(CASE WHEN NVL(FIICNT,0) > 1 THEN (NVL(FIICNT,0) - 1) "
    sql = sql & "          ELSE 0 END) DUPCNT " & vbCrLf
    sql = sql & " FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
    sql = sql & " WHERE FIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss') " & vbCrLf
    sql = sql & "   AND FIRKBN = " & mYimp.RecordIsTotal & vbCrLf           '//���v���R�[�h
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If Not dyn.EOF Then
        total(0) = dyn.Fields("OK_CNT")
        total(1) = dyn.Fields("NG_CNT")
        total(2) = dyn.Fields("DUPCNT")
    End If
    Call dyn.Close
    Set dyn = Nothing
    '//���׍s�̍��v�s�����핪
    sql = "SELECT" & vbCrLf
    sql = sql & " SUM(DECODE(NVL(FIOKFG, 0)," & mYimp.updNormal & ",  NVL(FIICNT,0),0)) OK_CNT," & vbCrLf
    sql = sql & " SUM(DECODE(NVL(FIOKFG, 0)," & mYimp.updNormal & ",0,NVL(FIICNT,0)  )) NG_CNT," & vbCrLf
    sql = sql & " SUM(CASE WHEN NVL(FIOKFG, 0) = " & mYimp.updNormal & " AND NVL(FIICNT,0) > 1 THEN (NVL(FIICNT,0) - 1) "
    sql = sql & "          ELSE 0 END) DUPCNT   " & vbCrLf
    sql = sql & " FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
    sql = sql & " WHERE FIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss') " & vbCrLf
    sql = sql & "   AND FIRKBN IN (" & vbCrLf
    sql = sql & "       SELECT FISEQN" & vbCrLf
    sql = sql & "       FROM " & mYimp.TfFurikaeImport & " b " & vbCrLf
    sql = sql & "       WHERE FIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss') " & vbCrLf
    sql = sql & "         AND FIRKBN = " & mYimp.RecordIsTotal & vbCrLf           '//���v���R�[�h
    sql = sql & "         AND FIOKFG = " & mYimp.updNormal & vbCrLf
    sql = sql & "   )" & vbCrLf
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If Not dyn.EOF Then
        Detail(0) = dyn.Fields("OK_CNT")
        Detail(1) = dyn.Fields("NG_CNT")
        Detail(2) = dyn.Fields("DUPCNT")
    End If
    Call dyn.Close
    Set dyn = Nothing
    '//���׍s�̍��v�s���ُ핪
    sql = "SELECT" & vbCrLf
    sql = sql & " SUM(DECODE(NVL(FIOKFG, 0)," & mYimp.updNormal & ",  NVL(FIICNT,0),0)) TT_CNT," & vbCrLf
    sql = sql & " SUM(DECODE(NVL(FIOKFG, 0)," & mYimp.updNormal & ",0,NVL(FIICNT,0)  )) DT_CNT," & vbCrLf
    sql = sql & " SUM(CASE WHEN NVL(FIICNT,0) > 1 THEN (NVL(FIICNT,0) - 1) "
    sql = sql & "          ELSE 0 END) DUPCNT " & vbCrLf
    sql = sql & " FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
    sql = sql & " WHERE FIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss') " & vbCrLf
    sql = sql & "   AND FIRKBN IN (" & vbCrLf
    sql = sql & "       SELECT FISEQN" & vbCrLf
    sql = sql & "       FROM " & mYimp.TfFurikaeImport & " b " & vbCrLf
    sql = sql & "       WHERE FIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss') " & vbCrLf
    sql = sql & "         AND FIRKBN = " & mYimp.RecordIsTotal & vbCrLf           '//���v���R�[�h
    sql = sql & "         AND FIOKFG <>" & mYimp.updNormal & vbCrLf
    sql = sql & "   )" & vbCrLf
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If Not dyn.EOF Then
        BadCnt(0) = gdDBS.Nz(dyn.Fields("TT_CNT"), 0)
        BadCnt(1) = gdDBS.Nz(dyn.Fields("DT_CNT"), 0)
        BadCnt(2) = gdDBS.Nz(dyn.Fields("DUPCNT"), 0)
    End If
    Call dyn.Close
    Set dyn = Nothing
    
    '//�}�X�^���f���ɂ�������������̂ŋ��ʉ�
    If pMoveTempRecords(Condition, cImportToYotei) < 0 Then
        GoTo cmdUpdate_ClickError:
    End If
    Call gdDBS.Database.CommitTrans
    
    pgrProgressBar.Max = pgrProgressBar.Max
    fraProgressBar.Visible = False
    
    '//�X�e�[�^�X�s�̐���E����
    stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "���f����"
    Call MsgBox("�}�X�^���f�Ώ� = [" & cboImpDate.Text & "]" & vbCrLf & vbCrLf & _
            recCnt & " �����}�X�^���f����܂���.(���ׁF����{���d���̌���)" & vbCrLf & vbCrLf & _
            "�����f���ꂽ�ڍד��e" & vbCrLf & _
            "���v�F���큁" & total(0) & " (���d���F" & total(2) & ")" & "�ُ큁" & total(1) & vbCrLf & _
            "���ׁF���큁" & Detail(0) & " (���d���F" & Detail(2) & ")" & "�ُ큁" & Detail(1) + BadCnt(0) + BadCnt(1) & vbCrLf & _
            "�@�捞������" & Detail(0) + Detail(1) + total(0) + total(1) + BadCnt(0) + BadCnt(1) _
            , vbOKOnly + vbInformation, mCaption)
    Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] �� " & recCnt & " ���̔��f���������܂����B(���ׁF����{���d���̌���)" & _
                            "�@�ڍ׌�����" & Detail(0) & " (���d���F" & Detail(2) & ")" & _
                            "�@�捞������" & Detail(0) + Detail(1) + total(0) + total(1) + BadCnt(0) + BadCnt(1) _
                        )
    
    '//���X�g���Đݒ�
    Call pMakeComboBox
    '//�{�^����߂�
    cmdUpdate.Caption = cBtnUpdate
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
    Exit Sub
cmdUpdate_ClickError:
    Call gdDBS.Database.Rollback
    If Err Then
        Dim errCode As Integer, errMsg As String
        If gdDBS.Database.LastServerErr Then
            errCode = gdDBS.Database.LastServerErr
            errMsg = gdDBS.Database.LastServerErrText
        Else
            errCode = Err
            errMsg = Error
        End If
        fraProgressBar.Visible = False
        '//�X�e�[�^�X�s�̐���E����
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�}�X�^���f�G���[(" & errCode & ")"
        Call gdDBS.AutoLogOut(mCaption, "�G���[�������������߃}�X�^���f�͒��~����܂����B(Error=" & errMsg & ")")
        Call MsgBox("�}�X�^���f�Ώ� = [" & cboImpDate.Text & "]" & vbCrLf & _
                    "�̓G���[�������������߃}�X�^���f�͒��~����܂����B" & vbCrLf & errMsg, _
                vbOKOnly + vbCritical, mCaption)
    Else
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�}�X�^���f���f"
        Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] �̃}�X�^���f�͒��~����܂����B")
    End If
    '//�{�^����߂�
    cmdUpdate.Caption = cBtnUpdate
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    cboFIITKB.Visible = False
    Call mForm.Init(Me, gdDBS)
    Call mSprTotal.Init(sprTotal)
    Call mSprDetail.Init(sprDetail)
    mSprTotal.OperationMode = OperationModeNormal    '//�ҏW����̂ŕW����
    mSprDetail.OperationMode = OperationModeNormal    '//�ҏW����̂ŕW����
    lblDetailCount.Caption = ""
    lblDetailKingaku.Caption = ""
    lblDetailCancel.Caption = ""
    
    Dim ix As Long, temp As String
    '///////////////////////////////////////////////////////////////
    '//Spread �̈ϑ��Җ��pWORK gdDBS �Ɋ֐����������̂ŗ��p
    Call gdDBS.SetItakushaComboBox(cboFIITKB)
    For ix = 0 To cboFIITKB.ListCount - 1
        temp = temp & cboFIITKB.List(ix) & vbTab
    Next ix
    Call mSprTotal.ComboBox(eSprTotal.eItakuName, temp)    '//�ϑ��Җ��̗�ɓ��e��ݒ�
    
    '//SprTotal �̗񒲐�
    mSprTotal.Locked(eSprTotal.eErrorStts, -1) = True    '//�ҏW���b�N
    mSprTotal.Locked(eSprTotal.eKeiyakuName, -1) = True  '//�ҏW���b�N
'//2006/04/26 �������E�񐔒ǉ��̗��ҏW���b�N
    mSprTotal.Locked(eSprTotal.eMochikomiBi, -1) = True     '//�ҏW���b�N
    mSprTotal.Locked(eSprTotal.eImportCnt, -1) = True      '//�ҏW���b�N
    With sprTotal
        'Call sprMeisai_LostFocus    '//ToolTip ��ݒ�
        If True <> mReg.Debuged Then
            .MaxCols = eSprTotal.eMaxCols
            '//�G���[�������̂ŕ\����(eUseCol)�ȍ~�͔�\���ɂ���
            For ix = eSprTotal.eUseCols To eSprTotal.eMaxCols
                .ColWidth(ix) = 0
            Next ix
            '//���בS�̂̏C���t���O���Z�b�g
        End If
        .Tag = mSprTotal.RowNonEdit
    End With
    '//SprDetail �̗񒲐�
    mSprDetail.Locked(eSprDetail.eErrorStts, -1) = True '//�ҏW���b�N
    mSprDetail.Locked(eSprDetail.eMasterKouza, -1) = True  '//�ҏW���b�N
'//2006/04/27 ���͍���
'//    mSprDetail.Locked(eSprDetail.eImportKouza, -1) = True  '//�ҏW���b�N
'//2006/04/26 �������E�񐔒ǉ��̗��ҏW���b�N
    mSprDetail.Locked(eSprDetail.eMochikomiBi, -1) = True     '//�ҏW���b�N
    mSprDetail.Locked(eSprDetail.eImportCnt, -1) = True      '//�ҏW���b�N
    With sprDetail
        '//�����\���͂O���A���v�N���b�N���ɕ\�������悤��...�B
        dbcImportDetail.RecordSource = ""
        .MaxRows = 0
        If True <> mReg.Debuged Then
            'Call sprMeisai_LostFocus    '//ToolTip ��ݒ�
            .MaxCols = eSprDetail.eMaxCols
            '//�G���[�������̂ŕ\����(eUseCol)�ȍ~�͔�\���ɂ���
            For ix = eSprDetail.eUseCols To eSprDetail.eMaxCols
                .ColWidth(ix) = 0
            Next ix
        End If
        '//���בS�̂̏C���t���O���Z�b�g
        .Tag = mSprDetail.RowNonEdit
    End With
    '//�X�e�[�^�X�s�̐���E����
    stbStatus.Panels.Item(stbStatus.Panels.Count).Text = ""
    pgrProgressBar.Left = 15
    pgrProgressBar.Top = 15
    pgrProgressBar.Height = 255
    pgrProgressBar.Width = 7035
    fraProgressBar.Height = pgrProgressBar.Height + 30
    fraProgressBar.Width = pgrProgressBar.Width + 30
    fraProgressBar.Visible = False
    cboSort.ListIndex = 0
    Call fraProgressBar.ZOrder(0)   '//�őO�ʂ�
    Call pMakeComboBox
'    txtFurikaebi.Text = mReg.FurikaeDataImport
End Sub

Private Sub Form_Resize()
    '//����ȏ㏬��������ƃR���g���[�����B���̂Ő��䂷��
    If Me.Height < 8100 Then
        Me.Height = 8100
    End If
    If Me.Width < 11300 Then
        Me.Width = 11300
    End If
    Call mForm.Resize
    fraProgressBar.Left = 1860
    fraProgressBar.Top = Me.Height - 970
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mAbort = True
    Set mForm = Nothing
    Set mReg = Nothing
    Set frmFurikaeYoteiImport = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub mnuEnd_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

'//�u���v�f�[�^�v�Z���P�ʂɃG���[�ӏ����J���[�\��
Private Sub pSpreadTotalSetErrorStatus(Optional vReset As Boolean = False)
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim ErrStts() As Variant, ix As Integer, cnt As Long
    Dim ms As New MouseClass
    Call ms.Start
'    eErrorStts = 1  '   FIERROR �G���[���e�F�ُ�A����A�x��
'    eItakuName      '           �ϑ��Җ�
'    eKeiyakuCode    '   FIKYCD  �_���
'    eKeiyakuName    '           �_��Җ�
'    eKyoshitsuNo    '   FIKSCD  �����ԍ�
'    ePageNumber     '   FIPGNO  ��
'    eFirukaeDate    '   FIFKDT  �U�֓�
'    eHenkoCount     '   FIHKCT  �ύX����
'    eHenkoKingaku   '   FIHKCT  �ύX���z
'    eCancelCount    '   FIKYCT  ��񌏐�
    
    If sprTotal.MaxRows = 0 Then
        Exit Sub
    End If
    '//�R�}���h�E�{�^������
    Call pLockedControl(False)
    '//�G���[���ݒ�
    ErrStts = Array("FIERROr", Empty, Empty, "FIITKBe", "FIKYCDe", "fikycde", "FIKSCDe", "FIPGNOe", "FIFKDTe", _
                    "FIHKCTe", "FIHKKGe", "FIKYCTe" _
                )
    sql = "SELECT ROWNUM,a.* FROM(" & vbCrLf
    sql = sql & "SELECT FIINDT,FISEQN," & mYimp.StatusColumns("," & vbCrLf, Len("," & vbCrLf))
    sql = sql & " FROM " & mYimp.TfFurikaeImport & " a " & vbCrLf
    sql = sql & " WHERE FIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss') " & vbCrLf
    sql = sql & "   AND FIRKBN = " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & " ORDER BY " & cSQLOrderString & vbCrLf
    '//�ȍ~�̂n�q�c�d�q��
    Select Case cboSort.ListIndex
    Case eSort.eImportSeq
        sql = sql & ",FIINDT,FISEQN" & vbCrLf
    Case eSort.eKeiyakusha
        sql = sql & ",FIITKB,FIKYCD,FIKSCD,FIPGNO,FIFKDT,FIHGCD,FISEQN" & vbCrLf
    Case Else
    End Select
    sql = sql & ") a"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If False = vReset Then
        'SPread �̃X�N���[���o�[�������̂݊J�n�s�Ɉړ�
        Call dyn.FindFirst("ROWNUM >= " & sprTotal.TopRow)
    End If
    mSprTotal.Redraw = False
    cnt = 0
    Do Until dyn.EOF
        '//���������G�ɂȂ�̂łƂɂ����Q�T�s�͓ǂ�ł��܂��I
        'If 0 = dyn.Fields(ErrStts(0)) And "����" = mSpread.Value(eSprCol.eErrorStts, dyn.RowPosition) Then
        '    Exit Do     '//�ُ�A�x���A����̃f�[�^���ɕ���ł���͂��Ȃ̂Ő���f�[�^�������Ȃ�I�����Ă��I
        'End If
        cnt = cnt + 1
        If cnt > cVisibleRows Then    '//���z���[�h�Ȃ̂łQ�T�s�ݒ肵�����_�ŏI��
            Exit Do
        End If
        For ix = LBound(ErrStts) To UBound(ErrStts)
            '//�e��̕\���F�ύX
            If Not IsEmpty(ErrStts(ix)) Then
                mSprTotal.BackColor(ix + 1, dyn.RowPosition) = mYimp.ErrorStatus(dyn.Fields(ErrStts(ix)))
            End If
        Next ix
        '//�������ʗ�̕\���F
        If mYimp.ErrorStatus(mYimp.errNormal) = mSprTotal.BackColor(eSprTotal.eErrorStts, dyn.RowPosition) Then
            mSprTotal.BackColor(eSprTotal.eErrorStts, dyn.RowPosition) = vbCyan
        End If
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    Set dyn = Nothing
    mSprTotal.Redraw = True
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
End Sub

'//�u���׃f�[�^�v�Z���P�ʂɃG���[�ӏ����J���[�\��
Private Sub pSpreadDetailSetErrorStatus(vImpDate As String, vSeqNo As Long, Optional vReset As Boolean = False)
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim ErrStts() As Variant, ix As Integer, cnt As Long
    Dim ms As New MouseClass
    Call ms.Start
'    eErrorStts = 1  '   FIERROR �G���[���e�F�ُ�A����A�x��
'    eHogoshaNo      '   FIHGCD  �ی�Ҕԍ�
'    eMasterKouza
'    eImportKouza
'    eHenkoGaku      '   FIHKKG  �ύX���z
'    eCancelFlag     '   FIKYFG  ���t���O
    
    If sprDetail.MaxRows = 0 Then
        Exit Sub
    End If
    '//�R�}���h�E�{�^������
    Call pLockedControl(False)
    '//�G���[���ݒ�
    ErrStts = Array("FIERROr", Empty, Empty, "FIHGCDe", "FIKZNMe", "fikznme", "FIHKKGe", "FIKYFGe" _
                )
    sql = "SELECT ROWNUM,a.* FROM(" & vbCrLf
    sql = sql & "SELECT TO_CHAR(FIINDT,'yyyy/mm/dd hh24:mi:ss') FIINDT,FISEQN," & mYimp.StatusColumns("," & vbCrLf, Len("," & vbCrLf))
    sql = sql & " FROM " & mYimp.TfFurikaeImport & " a "
    sql = sql & " WHERE (" & cInSQLString & ") IN(" & vbCrLf
    sql = sql & "       SELECT " & cInSQLString & vbCrLf
    sql = sql & "       FROM " & mYimp.TfFurikaeImport & " b " & vbCrLf
    sql = sql & "       WHERE FIINDT = TO_DATE(" & gdDBS.ColumnDataSet(cboImpDate.Text, vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss')" & vbCrLf
    sql = sql & "         AND FISEQN = " & vSeqNo & vbCrLf
    sql = sql & "         AND FIRKBN = " & mYimp.RecordIsTotal & vbCrLf
    sql = sql & "       )"
    sql = sql & "   AND FIRKBN <> " & mYimp.RecordIsTotal & vbCrLf
'//���ׂ͂r�d�p����
#If DETAIL_SEQN_ORDER = True Then
    sql = sql & " ORDER BY FIINDT,FISEQN" & vbCrLf
#Else
    sql = sql & " ORDER BY " & cSQLOrderString & vbCrLf
    '//�ȍ~�̂n�q�c�d�q��
    Select Case cboSort.ListIndex
    Case eSort.eImportSeq
        sql = sql & ",FIINDT,FISEQN" & vbCrLf
    Case eSort.eKeiyakusha
        sql = sql & ",FIITKB,FIKYCD,FIKSCD,FIPGNO,FIFKDT,FIHGCD,FISEQN" & vbCrLf
    Case Else
    End Select
#End If
    sql = sql & ") a"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If False = vReset Then
        'SPread �̃X�N���[���o�[�������̂݊J�n�s�Ɉړ�
        Call dyn.FindFirst("ROWNUM >= " & sprDetail.TopRow)
    End If
    mSprDetail.Redraw = False
    cnt = 0
    Do Until dyn.EOF
        '//���������G�ɂȂ�̂łƂɂ����Q�T�s�͓ǂ�ł��܂��I
        'If 0 = dyn.Fields(ErrStts(0)) And "����" = mSpread.Value(eSprCol.eErrorStts, dyn.RowPosition) Then
        '    Exit Do     '//�ُ�A�x���A����̃f�[�^���ɕ���ł���͂��Ȃ̂Ő���f�[�^�������Ȃ�I�����Ă��I
        'End If
        cnt = cnt + 1
        If cnt > cVisibleRows Then    '//���z���[�h�Ȃ̂łQ�T�s�ݒ肵�����_�ŏI��
            Exit Do
        End If
        For ix = LBound(ErrStts) To UBound(ErrStts)
            '//�e��̕\���F�ύX
            If Not IsEmpty(ErrStts(ix)) Then
                mSprDetail.BackColor(ix + 1, dyn.RowPosition) = mYimp.ErrorStatus(dyn.Fields(ErrStts(ix)))
            End If
        Next ix
        '//�������ʗ�̕\���F
        If mYimp.ErrorStatus(mYimp.errNormal) = mSprDetail.BackColor(eSprDetail.eErrorStts, dyn.RowPosition) Then
            mSprDetail.BackColor(eSprDetail.eErrorStts, dyn.RowPosition) = vbCyan
        End If
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    Set dyn = Nothing
    mSprDetail.Redraw = True
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
End Sub

Private Sub sprDetail_Change(ByVal Col As Long, ByVal Row As Long)
    '//�X�V���f�p�ɏC���t���O�ݒ�
    mSprDetail.Value(eSprDetail.eEditFlag, Row) = mSprDetail.RowEdit
    '//���׍s�ɕ�����ݒ�
    mSprDetail.Value(eSprDetail.eErrorStts, Row) = cEditDataMsg
    '//���׍s�̕�����F�ݒ�
    mSprDetail.BackColor(eSprDetail.eErrorStts, Row) = mYimp.ErrorStatus(mYimp.errEditData)
    '//���׍s�ɏC���t���O�ݒ�
    mSprDetail.Value(eSprDetail.eErrorFlag, Row) = mYimp.errEditData
    '//Tag �ɏC�������I���}�[�L���O
    sprDetail.Tag = mSprDetail.RowEdit
    cmdSprUpdate.Enabled = True
End Sub

Private Sub sprDetail_Click(ByVal Col As Long, ByVal Row As Long)
    '//���t���O�̃`�F�b�N�{�^������ => sprDetail_ButtonClicked() �ł���Ɩ��ׂ�\�����邽�тɖ��񔭐�����̂őʖځI
    Select Case Col
    Case eSprDetail.eCancelFlag
        Call sprDetail_Change(Col, Row)
    End Select
End Sub

Private Sub sprDetail_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    '// OldTop = 1 �̎��̓C�x���g���N���Ȃ�
#If True = VIRTUAL_MODE Then
    Call pSpreadDetailSetErrorStatus(cboImpDate.Text, mSprTotal.Value(eSprTotal.eImpSEQ, sprTotal.ActiveRow))
#Else
    If OldTop <> NewTop Then     '//���ׂăo�b�t�@�ɂ���̂őO�s�ɖ߂鎞�͂��Ȃ��悤��
        Call pSpreadDetailSetErrorStatus(cboImpDate.Text, mSprTotal.Value(eSprTotal.eImpSEQ, sprTotal.ActiveRow))
    End If
#End If
End Sub

Private Sub sprTotal_Change(ByVal Col As Long, ByVal Row As Long)
    If Col <= eSprTotal.eFirukaeDate Then
        '//�L�[�����C�������ꍇ
        mSprTotal.Value(eSprTotal.eEditFlag, Row) = mSprTotal.RowEditHeader
    
    ElseIf Val(mSprTotal.Value(eSprTotal.eEditFlag, Row)) <> mSprTotal.RowEditHeader Then
        '//�O��L�[�����C�����Ă��Ȃ����݂̂ŁA�L�[���ȊO���C�������ꍇ
        mSprTotal.Value(eSprTotal.eEditFlag, Row) = mSprTotal.RowEdit
    End If
    '//Tag �ɏC�������I���}�[�L���O
    sprTotal.Tag = mSprTotal.RowEdit
    '//���׍s�ɕ�����ݒ�
    mSprTotal.Value(eSprTotal.eErrorStts, Row) = cEditDataMsg
    '//���׍s�̕�����F�ݒ�
    mSprTotal.BackColor(eSprTotal.eErrorStts, Row) = mYimp.ErrorStatus(mYimp.errEditData)
    '//���׍s�ɏC���t���O�ݒ�
    mSprTotal.Value(eSprTotal.eErrorFlag, Row) = mYimp.errEditData
    cmdSprUpdate.Enabled = True
End Sub

Private Function pSpreadCheckAndUpdate(vMode As Boolean) As Boolean
    'If sprTotal.Tag = mSprTotal.RowEdit Or sprDetail.Tag = mSprDetail.RowEdit Then
    If True = vMode Then
        Select Case MsgBox("���e���ύX����Ă��܂��B" & vbCrLf & vbCrLf & _
                           "�X�V���܂����H", vbYesNoCancel + vbInformation, mCaption)
        Case vbYes
            Call cmdSprUpdate_Click
        Case vbNo
            '//�ύX���e��j��
            '//���v�����ׂ̏C���t���O�����Z�b�g
            sprDetail.Tag = mSprDetail.RowNonEdit
            sprTotal.Tag = mSprTotal.RowNonEdit
            Call cboImpDate_Click
        Case vbCancel
            pSpreadCheckAndUpdate = True '// LeaveCell() ���L�����Z��
            Exit Function
        End Select
    End If
    'cmdSprUpdate.Enabled = False
End Function

'//2006/06/16 ���v�s��A���ŏC�����ɍX�V�{�^���� Enabled=False �ɂȂ�ׁA�X�V�{�^����Ԃ�ǉ�
Private Function pSpreadDetailChange(Optional ByVal Row As Long = -1, Optional vButton As Boolean = False) As Boolean
    If True = pSpreadCheckAndUpdate(vButton Or sprDetail.Tag = mSprDetail.RowEdit) Then
        pSpreadDetailChange = True  '// LeaveCell() ���L�����Z��
        Exit Function
    End If
    If Row <= 0 Then
        '//�R�}���h�{�^���������� Row = -1 �ƂȂ�
        Exit Function
    End If
    Dim ms As New MouseClass
    Call ms.Start
    '//�f�[�^�ǂݍ��݁� Spread �ɐݒ蔽�f
    Call pReadDetailDataAndSetting(mSprTotal.Value(eSprTotal.eImpDate, Row), mSprTotal.Value(eSprTotal.eImpSEQ, Row))
    cmdSprUpdate.Enabled = False
    sprDetail.Tag = mSprDetail.RowNonEdit
    '// LeaveCell() �C�x���g�� Cancel �t���O��ԋp
    pSpreadDetailChange = False
End Function

Private Sub sprTotal_Click(ByVal Col As Long, ByVal Row As Long)
    '//�N�����̂P��ڂ̂� LeaveCell �C�x���g���������Ȃ��̂Ő���
    If False = mLeaveCellEvents And Row > 0 Then
        Call pSpreadDetailChange(Row, cmdSprUpdate.Enabled)
    End If
End Sub

Private Sub sprTotal_ComboCloseUp(ByVal Col As Long, ByVal Row As Long, ByVal SelChange As Integer)
    '//�B���R���{�{�b�N�X���g�p���ċ����I�ɃR�[�h�擾
    '//2006/06/13 �_��҇������̎��̈ϑ��Җ��I�����ɃG���[�ɂȂ�א���
    If "" <> Trim(mSprTotal.Text(eSprTotal.eItakuName, Row)) Then
        cboFIITKB.Text = mSprTotal.Text(eSprTotal.eItakuName, Row)
    End If
'// 'Z' �����݂���悤�ɂȂ����̂� Val() ������
'    If Val(mSprTotal.Value(eSprTotal.eItakuCode, Row)) <> cboFIITKB.ItemData(cboFIITKB.ListIndex) Then
    If mSprTotal.Value(eSprTotal.eItakuCode, Row) <> cboFIITKB.ItemData(cboFIITKB.ListIndex) Then
        mSprTotal.Value(eSprTotal.eItakuCode, Row) = cboFIITKB.ItemData(cboFIITKB.ListIndex)
    End If
End Sub

Private Sub sprTotal_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If Row <> NewRow And NewRow > 0 Then
        Cancel = pSpreadDetailChange(NewRow, cmdSprUpdate.Enabled)
        '//�N�����̂P��ڂ̂� LeaveCell �C�x���g���������Ȃ��̂Ő���
        mLeaveCellEvents = True
    End If
End Sub

Private Sub pSpreadRightClick(vRow As Long, vHeader As Boolean)
    mnuTitle.Caption = "�y�Ώہ�"
    If True = vHeader Then
        '//�폜�Ώۂ̂r�d�p�F�w�b�_���͖��ׂ������ɍ폜
        mDeleteSeqNo = mSprTotal.Value(eSprTotal.eImpSEQ, vRow)
        mnuTitle.Caption = mnuTitle.Caption & mSprTotal.Text(eSprTotal.eKeiyakuCode, vRow) & "-" & _
                                              mSprTotal.Text(eSprTotal.eKyoshitsuNo, vRow) & "-" & _
                                              mSprTotal.Text(eSprTotal.ePageNumber, vRow)
        mnuSprDelete.Caption = "���v�s��"
        mnuSprReset.Caption = "���v�s��"
        mnuSprDelete.Enabled = mSprTotal.Value(eSprTotal.eErrorFlag, vRow) <> mYimp.errDeleted
        mnuSprReset.Enabled = Not mnuSprDelete.Enabled
    Else
        '//�폜�Ώۂ̂r�d�p�F�w�b�_���͖��ׂ������ɍ폜
        mDeleteSeqNo = mSprTotal.Value(eSprDetail.eImpSEQ, vRow)
        mnuTitle.Caption = mnuTitle.Caption & mSprDetail.Text(eSprDetail.eKeiyakuCode, vRow) & "-" & _
                                              mSprDetail.Text(eSprDetail.eKyoshitsuNo, vRow) & "-" & _
                                              mSprDetail.Text(eSprDetail.ePageNumber, vRow) & "-" & _
                                              mSprDetail.Text(eSprDetail.eHogoshaNo, vRow)
        mnuSprDelete.Caption = ""
        mnuSprReset.Caption = ""
        mnuSprDelete.Enabled = mSprTotal.Value(eSprDetail.eErrorFlag, vRow) <> mYimp.errDeleted
        mnuSprReset.Enabled = Not mnuSprDelete.Enabled
    End If
    mnuTitle.Caption = mnuTitle.Caption & "�z"
    mnuSprDelete.Caption = mnuSprDelete.Caption & "���׍s���폜(&D)"
    mnuSprReset.Caption = mnuSprReset.Caption & "���׍s�̍폜������(&R)"
    mDeleteMenu = ePopup.eNoMenu '//�폜�A�N�V�����̃��j���[ -1=Delete,0=NonMenu,1=Reset
    Call PopupMenu(mnuSpread)
    Select Case mDeleteMenu
    Case ePopup.eNoMenu
    Case ePopup.eDelete
        If vHeader = True Then
            mSprTotal.Text(eSprTotal.eErrorStts, vRow) = cDeleteMsg
            mSprTotal.BackColor(eSprTotal.eErrorStts, vRow) = mYimp.ErrorStatus(mYimp.errDeleted)
        Else
            mSprDetail.Text(eSprDetail.eErrorStts, vRow) = cDeleteMsg
            mSprDetail.BackColor(eSprDetail.eErrorStts, vRow) = mYimp.ErrorStatus(mYimp.errDeleted)
        End If
    Case ePopup.eReset
        If vHeader = True Then
            mSprTotal.Text(eSprTotal.eErrorStts, vRow) = cEditDataMsg
            mSprTotal.BackColor(eSprTotal.eErrorStts, vRow) = mYimp.ErrorStatus(mYimp.errEditData)
        Else
            mSprDetail.Text(eSprDetail.eErrorStts, vRow) = cEditDataMsg
            mSprDetail.BackColor(eSprDetail.eErrorStts, vRow) = mYimp.ErrorStatus(mYimp.errEditData)
        End If
    End Select
End Sub

Private Sub mnuSprReset_Click()
    Dim sql As String, recCnt As Long
    sql = "UPDATE " & mYimp.TfFurikaeImport & " SET " & vbCrLf
    '//�C����Ԃ�
    sql = sql & " FIERROR = " & mYimp.errEditData & vbCrLf
    sql = sql & " WHERE FIINDT = TO_DATE(" & gdDBS.ColumnDataSet(cboImpDate.Text, vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss') " & vbCrLf
    sql = sql & "   AND(FIRKBN = " & gdDBS.ColumnDataSet(mDeleteSeqNo, "L", vEnd:=True) & vbCrLf
    sql = sql & "    OR FISEQN = " & gdDBS.ColumnDataSet(mDeleteSeqNo, "L", vEnd:=True) & vbCrLf
    sql = sql & "     )"
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    If recCnt Then
        mDeleteMenu = ePopup.eReset
    End If
End Sub

Private Sub mnuSprDelete_Click()
    Dim sql As String, recCnt As Long
    sql = "UPDATE " & mYimp.TfFurikaeImport & " SET " & vbCrLf
    sql = sql & " FIERROR = " & mYimp.errDeleted & "," & vbCrLf
'//�}�X�^���f�ΏۊO�ɂ���
    sql = sql & " FIOKFG  = " & mYimp.updInvalid & vbCrLf
    sql = sql & " WHERE FIINDT = TO_DATE(" & gdDBS.ColumnDataSet(cboImpDate.Text, vEnd:=True) & ",'yyyy/mm/dd hh24:mi:ss') " & vbCrLf
    sql = sql & "   AND(FIRKBN = " & gdDBS.ColumnDataSet(mDeleteSeqNo, "L", vEnd:=True) & vbCrLf
    sql = sql & "    OR FISEQN = " & gdDBS.ColumnDataSet(mDeleteSeqNo, "L", vEnd:=True) & vbCrLf
    sql = sql & "     )"
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    If recCnt Then
        mDeleteMenu = ePopup.eDelete
    End If
End Sub

Private Sub sprDetail_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If sprDetail.SelBlockRow = Row Then
        Call pSpreadRightClick(Row, False)
    End If
End Sub

Private Sub sprTotal_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If sprTotal.SelBlockRow = Row Then
        Call pSpreadRightClick(Row, True)
    End If
End Sub

Private Sub sprTotal_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    '// OldTop = 1 �̎��̓C�x���g���N���Ȃ�
#If True = VIRTUAL_MODE Then
    Call pSpreadTotalSetErrorStatus
#Else
    If OldTop <> NewTop Then     '//���ׂăo�b�t�@�ɂ���̂őO�s�ɖ߂鎞�͂��Ȃ��悤��
        Call pSpreadTotalSetErrorStatus
    End If
#End If
End Sub
#End If ' NO_RELEASE
