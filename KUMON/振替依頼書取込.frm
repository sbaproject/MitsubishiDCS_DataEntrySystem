VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmFurikaeReqImport 
   Caption         =   "�U�ֈ˗���(�捞)"
   ClientHeight    =   7965
   ClientLeft      =   2445
   ClientTop       =   2370
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11100
   Begin VB.CommandButton cmdDelete 
      Caption         =   "�p��(&D)"
      Height          =   435
      Left            =   5340
      TabIndex        =   7
      Top             =   6900
      Width           =   1395
   End
   Begin VB.Frame fraProgressBar 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "fraProgressBar"
      ForeColor       =   &H80000004&
      Height          =   290
      Left            =   1860
      TabIndex        =   12
      Top             =   7500
      Width           =   7060
      Begin MSComctlLib.ProgressBar pgrProgressBar 
         Height          =   255
         Left            =   15
         TabIndex        =   13
         Top             =   15
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
      End
   End
   Begin VB.ComboBox cboSort 
      Height          =   300
      ItemData        =   "�U�ֈ˗����捞.frx":0000
      Left            =   4500
      List            =   "�U�ֈ˗����捞.frx":000D
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   2
      Top             =   60
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�}�X�^���f(&U)"
      Height          =   435
      Left            =   6840
      TabIndex        =   8
      Top             =   6900
      Width           =   1395
   End
   Begin VB.CommandButton cmdErrList 
      Caption         =   "�G���[���X�g(&P)"
      Height          =   435
      Left            =   3480
      TabIndex        =   6
      Top             =   6900
      Width           =   1395
   End
   Begin FPSpread.vaSpread sprMeisai 
      Bindings        =   "�U�ֈ˗����捞.frx":0037
      Height          =   6315
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   420
      Width           =   10665
      _Version        =   196608
      _ExtentX        =   18812
      _ExtentY        =   11139
      _StockProps     =   64
      ButtonDrawMode  =   4
      ColsFrozen      =   6
      DAutoCellTypes  =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o����"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   51
      MaxRows         =   1000000
      SpreadDesigner  =   "�U�ֈ˗����捞.frx":004F
      UserResize      =   0
      VirtualMode     =   -1  'True
      VirtualScrollBuffer=   -1  'True
   End
   Begin VB.ComboBox cboImpDate 
      Height          =   300
      Left            =   1200
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   1
      Top             =   60
      Width           =   1935
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "�`�F�b�N(&C)"
      Height          =   435
      Left            =   1980
      TabIndex        =   5
      Top             =   6900
      Width           =   1395
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "�捞(&I)"
      Height          =   435
      Left            =   480
      TabIndex        =   4
      Top             =   6900
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "�I��(&X)"
      Height          =   435
      Left            =   9360
      TabIndex        =   0
      Top             =   6900
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   8580
      Top             =   6900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ORADCLibCtl.ORADC dbcImport 
      Height          =   315
      Left            =   9120
      Top             =   7320
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
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
      RecordSource    =   "select * from tchogoshaimport"
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  '������
      Height          =   315
      Left            =   0
      TabIndex        =   11
      Top             =   7650
      Width           =   11100
      _ExtentX        =   19579
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
   Begin VB.Label lblModoriCount 
      Caption         =   "�y �����߂茏���F 9,999 �� �z"
      BeginProperty Font 
         Name            =   "�l�r �o����"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6540
      TabIndex        =   15
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "�\����"
      Height          =   180
      Left            =   3780
      TabIndex        =   14
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label8 
      Caption         =   "�捞����"
      Height          =   180
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   780
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   195
      Left            =   9540
      TabIndex        =   9
      Top             =   0
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
Attribute VB_Name = "frmFurikaeReqImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//���z���[�h�ł���Ɠ������ς��I�I�I
#Const VIRTUAL_MODE = True
#Const DATA_DUPLICATE = False   '//�U�ֈ˗����͏d�����`�F�b�N

#Const BLOCK_CHECK = False           '//�`�F�b�N���̃u���b�N���������邩�H��\���F�f�o�b�N���̂�
#If BLOCK_CHECK = True Then           '//�`�F�b�N���̃u���b�N���������邩�H��\���F�f�o�b�N���̂�
Private mCheckBlocks As Integer
#End If

Private mCaption As String
Private mAbort As Boolean
Private mForm As New FormClass
Private mSpread As New SpreadClass
Private mReg As New RegistryClass
Private mRimp As New FurikaeReqImpClass
Public mEditRow As Long      '//�C�����̍s�ԍ�

Private Type tpErrorStatus
    Field   As String
    Error   As Integer
    Message As String
End Type
Private mErrStts() As tpErrorStatus

Private Type tpHogoshaImport
    MochikomiBi     As String * 8   '//�������ݓ� 2006/03/24 ���ڒǉ�
    Keiyakusha      As String * 5   '//�_��Ҕԍ�
    Kyoshittsu      As String * 1   '//�����ԍ�
    HogoshaNo       As String * 4   '//�ی�Ҕԍ�
    HogoshaKana     As String * 40  '//�ی�Җ�(�J�i)=>�������`�l
    HogoshaKanji    As String * 30  '//�ی�Җ�(����)
    SeitoShimei     As String * 50  '//���k����
    BankCode        As String * 4   '//���Z�@�փR�[�h
    BankName        As String * 30  '//���Z�@�֖�
    ShitenCode      As String * 3   '//�x�X�R�[�h
    ShitenName      As String * 30  '//�x�X��
    YokinShumoku    As String * 1   '//�a�����
    KouzaBango      As String * 7   '//�����ԍ�
    TuuchoKigou     As String * 3   '//�ʒ��L��
    TuuchoBango     As String * 8   '//�ʒ��ԍ�
    FurikaeGaku     As String * 6   '//�U�֋��z
    CrLf            As String * 2   '// CR + LF
End Type

Private Const cBtnCancel As String = "���~(&A)"
Private Const cBtnImport As String = "�捞(&I)"
Private Const cBtnDelete As String = "�p��(&D)"
Private Const cBtnCheck  As String = "�`�F�b�N(&C)"
Private Const cBtnUpdate As String = "�}�X�^���f(&U)"
Private Const cVisibleRows As Long = 25
'Private Const cImportToUpdate As String = "U"
'Private Const cImportToInsert As String = "I"
Private Const cEditDataMsg  As String = "�C�� => �`�F�b�N���������ĉ������B"
Private Const cImportMsg    As String = "�捞 => �`�F�b�N���������ĉ������B"

Private Enum eSprCol
    eErrorStts = 1  '�G���[���e�F�ُ�A����A�x��
    eItakuName      '   CIITKB  �ϑ��Җ�
    eKeiyakuCode    '   CIKYCD  �_��҃R�[�h
    eKeiyakuName    '           �_��Җ�
    eKyoshitsuNo    '   CIKSCD  �����ԍ�
    eHogoshaCode    '   CIHGCD  �ی�҃R�[�h
    eHogoshaName    '   CIKJNM  �ی�Җ�(����)
    eHogoshaKana    '   CIKNNM  �ی�Җ�(�J�i)=>�������`�l��
    eSeitoName      '   CISTNM  ���k����
    eFurikaeGaku    '   CISKGK  �U�֋��z
    eKinyuuKubun    '   CIKKBN  ���Z�@�֋敪
    eBankCode       '   CIBANK  ��s�R�[�h
    eBankName_m     '           ��s��(�}�X�^�[)
    eBankName_i     '   CIBKNM  ��s��(�捞)
    eShitenCode     '   CISITN  �x�X�R�[�h
    eShitenName_m   '           �x�X��(�}�X�^�[)
    eShitenName_i   '   CISINM  �x�X��(�捞)
    eYokinShumoku   '   CIKZSB  �a�����
    eKouzaBango     '   CIKZNO  �����ԍ�
    eYubinKigou     '   CIYBTK  �X�֋�:�ʒ��L��
    eYubinBango     '   CIYBTN  �X�֋�:�ʒ��ԍ�
    eKouzaName      '   CIKZNM  �������`�l=>�ی�Җ�(�J�i)
    eMstUpdate      '//�}�X�^�[���f�t���O
    eImpDate        '�捞��
    eImpSEQ         '�r�d�p
    eUseCols = eKouzaName  '//�\�������͍����܂�
    eMaxCols = 50   '//�G���[����܂߂āI
End Enum

Private Enum eSort
    eImportSeq
    eKeiyakusha
    eKinnyuKikan
End Enum
Private mMainSQL As String

Private Sub pLockedControl(blMode As Boolean, Optional vButton As CommandButton = Nothing)
    cmdImport.Enabled = blMode
    cmdCheck.Enabled = blMode
    cmdErrList.Enabled = blMode
    cmdDelete.Enabled = blMode
    cmdUpdate.Enabled = blMode
    cmdEnd.Enabled = blMode     '//�����r���ŏI������Ƃ��������Ȃ�̂ŏI�����E���I
    If Not vButton Is Nothing Then
        vButton.Enabled = True
    End If
End Sub

Private Function pMakeSQLReadData(Optional vErrColomns As Boolean = False) As String
    Dim sql As String
    
    sql = "SELECT * FROM(" & vbCrLf
    sql = sql & "SELECT " & vbCrLf
    'sql = sql & " CIERROR," & vbCrLf
#If SHORT_MSG Then
    sql = sql & " DECODE(CIERROR,-3,'�捞',-2,'�C��',-1,decode(cimupd,1,'�x��','�ُ�'),0,'����',1,'�x��','��O') as CIERRNM," & vbCrLf
#Else
    sql = sql & " CASE WHEN CIERROR = -2 THEN " & gdDBS.ColumnDataSet(cEditDataMsg, vEnd:=True) & vbCrLf
    sql = sql & "      WHEN CIERROR = -3 THEN " & gdDBS.ColumnDataSet(cImportMsg, vEnd:=True) & vbCrLf
    sql = sql & "      WHEN CIERROR IN(-1,+0,+1) THEN " & vbCrLf
    sql = sql & "           DECODE(CIERROR," & vbCrLf
    sql = sql & "               -1,decode(cimupd,1,'�x��','�ُ�')," & vbCrLf
    sql = sql & "               +0,'����'," & vbCrLf
    sql = sql & "               +1,'�x��'," & vbCrLf
    sql = sql & "               NULL" & vbCrLf
    sql = sql & "           ) || ' => ' || " & vbCrLf
    sql = sql & "       DECODE(CIOKFG," & vbCrLf
    sql = sql & "               " & mRimp.updInvalid & ",'" & mRimp.mUpdateMessage(mRimp.updInvalid) & "'," & vbCrLf
    sql = sql & "               " & mRimp.updWarnErr & ",'" & mRimp.mUpdateMessage(mRimp.updWarnErr) & "'," & vbCrLf
    sql = sql & "               " & mRimp.updNormal & ",'" & mRimp.mUpdateMessage(mRimp.updNormal) & "'," & vbCrLf
    sql = sql & "               " & mRimp.updWarnUpd & ",'" & mRimp.mUpdateMessage(mRimp.updWarnUpd) & "'," & vbCrLf
    sql = sql & "               " & mRimp.updResetCancel & ",'" & mRimp.mUpdateMessage(mRimp.updResetCancel) & "'," & vbCrLf
    sql = sql & "               '�������ʂ�����ł��܂���B'" & vbCrLf
    sql = sql & "           )" & vbCrLf
    sql = sql & "      ELSE                             '��O => �������ʂ�����ł��܂���B'" & vbCrLf
    sql = sql & " END as CIERRNM," & vbCrLf
#End If
    'sql = sql & " CIITKB," & vbCrLf
    sql = sql & " (SELECT ABKJNM " & vbCrLf
    sql = sql & "  FROM taItakushaMaster" & vbCrLf
    sql = sql & "  WHERE ABITKB = a.CIITKB" & vbCrLf
    sql = sql & " ) as ABKJNM," & vbCrLf    '//�ʏ�̊O�������ł���Ƃ�₱�����̂�...(tcHogoshaImport Table �͑S���o�������I)
    sql = sql & " CIKYCD," & vbCrLf
    sql = sql & " (SELECT MAX(BAKJNM) BAKJNM " & vbCrLf
    sql = sql & "  FROM tbKeiyakushaMaster " & vbCrLf
    sql = sql & "  WHERE BAITKB = a.CIITKB" & vbCrLf
    sql = sql & "    AND BAKYCD = a.CIKYCD" & vbCrLf
    '//�_��҂͌��ݗL�����F�_����ԁ��U�֊���
'//2012/08/09 �_����Ԃ𕜊��F�Â��������o�Ă��܂��o�O�Ή�
'    sql = sql & "    AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN BAKYST AND BAKYED" & vbCrLf
'    sql = sql & "    AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN BAFKST AND BAFKED" & vbCrLf
    sql = sql & "     and basqno in(" & vbCrLf
    sql = sql & "       select max(basqno) from tbKeiyakushaMaster " & vbCrLf
    sql = sql & "       WHERE BAITKB = a.CIITKB" & vbCrLf
    sql = sql & "         AND BAKYCD = a.CIKYCD" & vbCrLf
    sql = sql & "   )"
    sql = sql & " ) as BAKJNM," & vbCrLf    '//�ʏ�̊O�������ł���Ƃ�₱�����̂�...(tcHogoshaImport Table �͑S���o�������I)
    sql = sql & " CIKSCD," & vbCrLf
    sql = sql & " CIHGCD," & vbCrLf
    sql = sql & " CIKJNM," & vbCrLf
    sql = sql & " CIKNNM," & vbCrLf
    sql = sql & " CISTNM," & vbCrLf
    sql = sql & " DECODE(CISKGK,NULL,'',TO_CHAR(CISKGK,'99,999,999')) as CISKGK," & vbCrLf
    sql = sql & " DECODE(CIKKBN," & eBankKubun.KinnyuuKikan & ",'����'," & eBankKubun.YuubinKyoku & ",'�X�֋�',NULL)     as CIKKBN," & vbCrLf
    sql = sql & " CIBANK," & vbCrLf
    sql = sql & " (SELECT DAKJNM" & vbCrLf
    sql = sql & "  FROM tdBankMaster" & vbCrLf
    sql = sql & "  WHERE DABANK = a.CIBANK" & vbCrLf
    sql = sql & "    AND DARKBN = '0'"
    sql = sql & "    AND DASITN = '000'"
    sql = sql & "    AND DASQNO = ':'"      '//���ꂪ���ݗL��
    sql = sql & " ) as DABKNM,"                '//�ʏ�̊O�������ł���Ƃ�₱�����̂�...(tcHogoshaImport Table �͑S���o�������I)
    sql = sql & " CIBKNM," & vbCrLf
    sql = sql & " CISITN," & vbCrLf
    sql = sql & " (SELECT DAKJNM" & vbCrLf
    sql = sql & "  FROM tdBankMaster" & vbCrLf
    sql = sql & "  WHERE DABANK = a.CIBANK" & vbCrLf
    sql = sql & "    AND DASITN = a.CISITN"
    sql = sql & "    AND DARKBN = '1'"
    sql = sql & "    AND DASQNO = '�'"      '//���ꂪ���ݗL��
    sql = sql & " ) as DASTNM,"                '//�ʏ�̊O�������ł���Ƃ�₱�����̂�...(tcHogoshaImport Table �͑S���o�������I)
    sql = sql & " CISINM," & vbCrLf
    sql = sql & " DECODE(CIKKBN," & eBankKubun.KinnyuuKikan & ",DECODE(CIKZSB,'1','����','2','����',CIKZSB),NULL) as CIKZSB," & vbCrLf
    sql = sql & " CIKZNO," & vbCrLf
    sql = sql & " CIYBTK," & vbCrLf
    sql = sql & " CIYBTN," & vbCrLf
    sql = sql & " CIKZNM," & vbCrLf
    sql = sql & " CIMUPD," & vbCrLf     '//2006/04/04 �}�X�^���f�n�j�t���O���ڒǉ�
    sql = sql & " TO_CHAR(CIINDT,'yyyy/mm/dd hh24:mi:ss') CIINDT," & vbCrLf
    If vErrColomns Then
        sql = sql & mRimp.StatusColumns("," & vbCrLf)
    End If
    sql = sql & " CISEQN " & vbCrLf
    '////////////////////////////////////////////////////////////////////
    '//����ȍ~�̂r�p�k (MainSQL) ���C����ʂŗ��p����̂Œ��ӂ��ĕύX�̂��ƁI�I�I
    mMainSQL = " FROM " & mRimp.TcHogoshaImport & " a" & vbCrLf
    mMainSQL = mMainSQL & " WHERE CIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf
    '//2006/04/14 ORDER ���v�f�ʂ�ɂȂ��Ă��Ȃ�����
    'mMainSQL = mMainSQL & " ORDER BY DECODE(CIERSR,-2, 1,-1,-12, 1,-11 ,CIERSR)"    '�C���A�G���[�A�x���A����̏�
    mMainSQL = mMainSQL & " ORDER BY DECODE(CIERSR,-2, -11, -1,-12, 1,-10 ,CIERSR)"    '�C���A�G���[�A�x���A����̏�
    '//�ȍ~�̂n�q�c�d�q��
    Select Case cboSort.ListIndex
    Case eSort.eImportSeq
        mMainSQL = mMainSQL & ",CIINDT,CISEQN" & vbCrLf
    Case eSort.eKeiyakusha
        mMainSQL = mMainSQL & ",CIITKB,CIKYCD,CIKSCD,CIHGCD,CISEQN" & vbCrLf
    Case eSort.eKinnyuKikan
        mMainSQL = mMainSQL & ",CIKKBN,CIBANK,CISITN,CIKZSB,CIKZNO,CIYBTK,CIYBTN,CISEQN" & vbCrLf
    Case Else
    End Select
    sql = sql & mMainSQL & ")"
    pMakeSQLReadData = sql
End Function

Private Sub pReadDataAndSetting()
    
    dbcImport.RecordSource = pMakeSQLReadData
    sprMeisai.VirtualMode = False   '//��U���z���[�h����
    Call dbcImport.Refresh

#If True = VIRTUAL_MODE Then
    '//���z���[�h�ɂ���ƃy�[�W���ς��ƃf�[�^������ւ���Ă��܂��̂Œ��ӁI�I�I
    sprMeisai.VScrollSpecial = True
    sprMeisai.VScrollSpecialType = 0
    sprMeisai.VirtualMode = True    '//���z���[�h�Đݒ�F�s�̃��t���b�V���I
    '//2012/07/02 ����̃f�[�^�ɑ΂��ĕ\�����ł��Ȃ��H�o�O�H�Ȃ̂Őݒ�s���R�����g���FSQL�����������H
    sprMeisai.VirtualMaxRows = dbcImport.Recordset.RecordCount
#Else
    sprMeisai.VScrollSpecial = True
    sprMeisai.VScrollSpecialType = 0
    sprMeisai.MaxRows = dbcImport.Recordset.RecordCount
#End If
    
    '//�Z���P�ʂɃG���[�ӏ����J���[�\��
    Call pSpreadSetErrorStatus(True)
    '//ToolTip ��L���ɂ���ׂɋ����I�Ƀt�H�[�J�X���ڂ�
    'Call sprMeisai.SetFocus
'//2007/07/19 �����߂�̌�����\��
    Dim sql As String, dyn As OraDynaset
    sql = "select count(*) modori " & vbCrLf
    sql = sql & " from " & mRimp.TcHogoshaImport & " a," & vbCrLf
    sql = sql & "   (select " & vbCrLf
    sql = sql & "     distinct caitkb,cakycd,cakscd,cahgcd " & vbCrLf
    sql = sql & "     from tcHogoshaMaster" & vbCrLf
    sql = sql & "   ) b " & vbCrLf
    sql = sql & " where ciitkb = caitkb " & vbCrLf
    sql = sql & "   and cikycd = cakycd " & vbCrLf
    sql = sql & "   and cikscd = cakscd " & vbCrLf
    sql = sql & "   and cihgcd = cahgcd " & vbCrLf
    sql = sql & "   AND CIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss')"
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
    lblModoriCount.Caption = "�y �����߂茏���F " & Format(dyn.Fields("modori"), "#,0") & " �� �z"
    Call dyn.Close
    Set dyn = Nothing
End Sub

Private Function pCheckSubForm() As Boolean
    '//�C����ʂ��\������Ă����Ȃ���Ă��܂��I
    If Not gdFormSub Is Nothing Then
        '//�����Ȃ��H
        'If gdFormSub.dbcImport.EditMode <> OracleConstantModule.ORADATA_EDITNONE Then
            If vbOK <> MsgBox("�C����ʂł̌��ݕҏW���̃f�[�^�͔j������܂�." & vbCrLf & vbCrLf & "��낵���ł����H", vbOKCancel + vbDefaultButton2 + vbInformation, mCaption) Then
                Exit Function
            End If
            'Call gdFormSub.dbcImport.UpdateControls   '//�L�����Z��
            Call gdFormSub.cmdEnd_Click
        'End If
        'Unload gdFormSub
        Set gdFormSub = Nothing
    End If
    pCheckSubForm = True
End Function
Private Sub cboImpDate_Click()
    If "" = Trim(cboImpDate.Text) Then
        '//�L�蓾�Ȃ�
        Exit Sub
    End If
    If False = pCheckSubForm Then
        Exit Sub
    End If
    Dim ms As New MouseClass
    Call ms.Start
    '//�f�[�^�ǂݍ��݁� Spread �ɐݒ蔽�f
    Call pReadDataAndSetting
End Sub

Private Sub cboSort_Click()
    Call cboImpDate_Click
End Sub

Private Function pMoveTempRecords(vCondition As String, vMode As String) As Long
    Dim sql As String
    '//�폜�Ώۃf�[�^�� Temp �Ƀo�b�N�A�b�v
    sql = "INSERT INTO " & mRimp.TcHogoshaImport & "Temp" & vbCrLf
    sql = sql & " SELECT SYSDATE,'" & vMode & "',a.*"
    sql = sql & " FROM " & mRimp.TcHogoshaImport & " a " & vbCrLf
    sql = sql & " WHERE 1 = 1" & vbCrLf
    sql = sql & vCondition
    Call gdDBS.Database.ExecuteSQL(sql)
    
    sql = "DELETE " & mRimp.TcHogoshaImport & vbCrLf
    sql = sql & " WHERE 1 = 1" & vbCrLf
    sql = sql & vCondition
    pMoveTempRecords = gdDBS.Database.ExecuteSQL(sql)
End Function

Private Sub cmdDelete_Click()
    If False = pCheckSubForm Then
        Exit Sub
    End If
    If 0 = cboImpDate.ListCount Then
        Exit Sub
    ElseIf vbOK <> MsgBox("���ݕ\������Ă���f�[�^��j�����܂�." & vbCrLf & vbCrLf & _
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
    recCnt = pMoveTempRecords(" AND CIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss')", gcFurikaeImportToDelete)
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
    If err Then
        Dim errCode As Integer, errMsg As String
        If gdDBS.Database.LastServerErr Then
            errCode = gdDBS.Database.LastServerErr
            errMsg = gdDBS.Database.LastServerErrText
        Else
            errCode = err
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
    Unload Me
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

Private Sub cmdErrList_Click()
    If False = pCheckSubForm Then
        Exit Sub
    End If
    Dim reg As New RegistryClass
    Dim sql As String
    Load rptFurikaeReqImport
    With rptFurikaeReqImport
        .lblSort.Caption = "�\�����F " & cboSort.Text
        '.mTotalCnt = dbcImport.Recordset.RecordCount
        .documentName = mCaption
        .adoData.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=" & reg.DbPassword & _
                                    ";Persist Security Info=True;User ID=" & reg.DbUserName & _
                                                           ";Data Source=" & reg.DbDatabaseName
        sql = pMakeSQLReadData(True)
        '//�G���[�f�[�^�͈���ŏo�͂��Ȃ�
        sql = sql & " WHERE CIERROR <> " & mRimp.errNormal
        .adoData.Source = sql
        'Call .adoData.Refresh
        Call .Show
    End With
End Sub

Private Sub cmdImport_Click()
    If False = pCheckSubForm Then
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
    '//�U���˗����f�[�^���C���|�[�g
    Dim Hogosha As tpHogoshaImport
    Dim fp As Integer
    Dim ms As New MouseClass
    Call ms.Start
    
    fp = FreeFile
    Open dlgFile.FileName For Random Access Read As #fp Len = Len(Hogosha)
    fraProgressBar.Visible = True
    pgrProgressBar.Max = LOF(fp) / Len(Hogosha)
    '//�t�@�C���T�C�Y���Ⴄ�ꍇ�̌x�����b�Z�[�W
    If pgrProgressBar.Max <> Int(pgrProgressBar.Max) Then
        If (LOF(fp) - 1) / Len(Hogosha) <> Int((LOF(fp) - 1) / Len(Hogosha)) Then
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
    Dim insCnt As Long, updCnt As Long, recCnt As Long
    
    insDate = gdDBS.sysDate()
    
    Call gdDBS.Database.BeginTrans
    '///////////////////////////////////////////////
    '//�V�[�P���X���P�Ԃ���Ƀ��Z�b�g
    sql = "declare begin ResetSequence('sqImportSeq',1); end;"
    Call gdDBS.Database.ExecuteSQL(sql)
    
    Do While Loc(fp) < LOF(fp) / Len(Hogosha)
        DoEvents
        If mAbort Then
            GoTo cmdImport_ClickError
        End If
        Get #fp, , Hogosha
        recCnt = Loc(fp)
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = _
            "�c��" & Right(String(7, " ") & Format(pgrProgressBar.Max - Loc(fp), "#,##0"), 7) & " ��"
        pgrProgressBar.Value = IIf(Loc(fp) <= pgrProgressBar.Max, Loc(fp), pgrProgressBar.Max)
        
#If DATA_DUPLICATE = True Then  '//�U�ֈ˗����͏d�����`�F�b�N
'''''''''''''        '//2006/03/24 �ォ�玝�����񂾃f�[�^(���������傫��)���L���ɂ���
'''''''''''''        sql = "SELECT CIMCDT"
'''''''''''''        sql = sql & " FROM " & mRimp.TcHogoshaImport & " a "
'''''''''''''        sql = sql & "WHERE CIINDT = " & "TO_DATE(" & gdDBS.ColumnDataSet(insDate,"D", vEnd:=True) & ",'yyyy-mm-dd hh24:mi:ss')" & vbCrLf  '//�捞��
'''''''''''''        sql = sql & "  AND CIKYCD = " & gdDBS.ColumnDataSet(Hogosha.Keiyakusha, vEnd:=True) & vbCrLf    '//�_��Ҕԍ�
'''''''''''''        sql = sql & "  AND CIKSCD = " & gdDBS.ColumnDataSet(Hogosha.Kyoshittsu, vEnd:=True) & vbCrLf    '//�����ԍ�
'''''''''''''        sql = sql & "  AND CIHGCD = " & gdDBS.ColumnDataSet(Hogosha.HogoshaNo, vEnd:=True) & vbCrLf     '//�ی�Ҕԍ�
'''''''''''''#If ORA_DEBUG = 1 Then
'''''''''''''        Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
'''''''''''''#Else
'''''''''''''        Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
'''''''''''''#End If
'''''''''''''        If Not dyn.EOF Then
'''''''''''''            '//������������Ȃ�X�V
'''''''''''''            If Val(dyn.Fields("CIMCDT")) < Val(Hogosha.MochikomiBi) Then
'''''''''''''                '//�X�V�����݂�F����e�L�X�g���ɓ����f�[�^������H
'''''''''''''                sql = "UPDATE " & mRimp.TcHogoshaImport & " SET " & vbCrLf
'''''''''''''                '//�ϑ��҂͌_��҂��犄��o���̂ŕs�v�I
'''''''''''''                'sql = sql & "CIITKB = (SELECT ABITKB FROM taItakushaMaster WHERE ABKYTP = '" & Left(Hogosha.Keiyakusha, 1) & "')," & vbcrlf  '//�ϑ��ҋ敪
'''''''''''''                sql = sql & "CIKJNM = " & gdDBS.ColumnDataSet(Hogosha.HogoshaKanji) & vbCrLf    '//�ی�Җ�_����
'''''''''''''                sql = sql & "CIKNNM = " & gdDBS.ColumnDataSet(Hogosha.HogoshaKana) & vbCrLf     '//�ی�Җ�_�J�i
'''''''''''''                sql = sql & "CISTNM = " & gdDBS.ColumnDataSet(Hogosha.SeitoShimei) & vbCrLf     '//���k����
'''''''''''''                sql = sql & "CISKGK = " & gdDBS.ColumnDataSet(Hogosha.FurikaeGaku, "L") & vbCrLf '//�U�֋��z
'''''''''''''                sql = sql & "CIBKNM = " & gdDBS.ColumnDataSet(Hogosha.BankName) & vbCrLf        '//�捞��s��
'''''''''''''                sql = sql & "CISINM = " & gdDBS.ColumnDataSet(Hogosha.ShitenName) & vbCrLf      '//�捞�x�X��
'''''''''''''                If "" = Trim(Hogosha.TuuchoKigou) _
'''''''''''''                And "" = Trim(Hogosha.TuuchoBango) Then         '//�X�֋Ǐ��L���Ȃ�
'''''''''''''                    sql = sql & "CIKKBN = " & gdDBS.ColumnDataSet(eBankKubun.KinnyuuKikan, "I") & vbCrLf               '//������Z�@�֋敪
'''''''''''''                Else
'''''''''''''                    sql = sql & "CIKKBN = " & gdDBS.ColumnDataSet(eBankKubun.YuubinKyoku, "I") & vbCrLf               '//������Z�@�֋敪
'''''''''''''                End If
'''''''''''''                sql = sql & "CIBANK = " & gdDBS.ColumnDataSet(Hogosha.BankCode) & vbCrLf        '//�����s
'''''''''''''                sql = sql & "CISITN = " & gdDBS.ColumnDataSet(Hogosha.ShitenCode) & vbCrLf      '//����x�X
'''''''''''''                sql = sql & "CIKZSB = " & gdDBS.ColumnDataSet(Hogosha.YokinShumoku) & vbCrLf    '//�������
'''''''''''''                sql = sql & "CIKZNO = " & gdDBS.ColumnDataSet(Hogosha.KouzaBango) & vbCrLf      '//�����ԍ�
'''''''''''''                sql = sql & "CIYBTK = " & gdDBS.ColumnDataSet(Hogosha.TuuchoKigou) & vbCrLf     '//�ʒ��L��
'''''''''''''                sql = sql & "CIYBTN = " & gdDBS.ColumnDataSet(Hogosha.TuuchoBango) & vbCrLf     '//�ʒ��ԍ�
'''''''''''''                sql = sql & "CIKZNM = " & gdDBS.ColumnDataSet(Hogosha.HogoshaKana) & vbCrLf     '//�������`�l_�J�i
'''''''''''''                sql = sql & "CIERROR = " & gdDBS.ColumnDataSet(mRimp.errImport) & vbCrLf
'''''''''''''                sql = sql & "CIERSR  = " & gdDBS.ColumnDataSet(mRimp.errImport) & vbCrLf
'''''''''''''                sql = sql & "CIMCDT = " & gdDBS.ColumnDataSet(Hogosha.MochikomiBi, "L") & vbCrLf    '//������
'''''''''''''                sql = sql & "CIUPDT = SYSDATE " & vbCrLf                                        '//�X�V��
'''''''''''''                sql = sql & "WHERE CIINDT = " & "TO_DATE(" & gdDBS.ColumnDataSet(insDate,"D", vEnd:=True) & ",'yyyy-mm-dd hh24:mi:ss')" & vbCrLf  '//�捞��
'''''''''''''                sql = sql & "  AND CIKYCD = " & gdDBS.ColumnDataSet(Hogosha.Keiyakusha, vEnd:=True) & vbCrLf    '//�_��Ҕԍ�
'''''''''''''                sql = sql & "  AND CIKSCD = " & gdDBS.ColumnDataSet(Hogosha.Kyoshittsu, vEnd:=True) & vbCrLf    '//�����ԍ�
'''''''''''''                sql = sql & "  AND CIHGCD = " & gdDBS.ColumnDataSet(Hogosha.HogoshaNo, vEnd:=True) & vbCrLf     '//�ی�Ҕԍ�
'''''''''''''                Call gdDBS.Database.ExecuteSQL(sql)
'''''''''''''                updCnt = updCnt + 1&
'''''''''''''            End If
'''''''''''''        Else
#End If     '//#If DATA_DUPLICATE = True Then  '//�U�ֈ˗����͏d�����`�F�b�N
            
            insCnt = insCnt + 1&
            '//�X�V�ł��Ȃ������̂ő}�������݂�
            '//�f�[�^���e�[�u���ɑ}��
            sql = "INSERT INTO " & mRimp.TcHogoshaImport & "(" & vbCrLf
            sql = sql & "CIINDT,"   '//�捞��
            sql = sql & "CISEQN,"   '//�捞SEQNO
            sql = sql & "CIITKB,"   '//�ϑ��ҋ敪
            sql = sql & "CIKYCD,"   '//�_��Ҕԍ�
            sql = sql & "CIKSCD,"   '//�����ԍ�
            sql = sql & "CIHGCD,"   '//�ی�Ҕԍ�
            sql = sql & "CIKJNM,"   '//�ی�Җ�_����
            sql = sql & "CIKNNM,"   '//�ی�Җ�_�J�i
            sql = sql & "CISTNM,"   '//���k����
            sql = sql & "CISKGK,"   '//�U�֋��z
            sql = sql & "CIBKNM,"   '//�捞��s��
            sql = sql & "CISINM,"   '//�捞�x�X��
            sql = sql & "CIKKBN,"   '//������Z�@�֋敪
            sql = sql & "CIBANK,"   '//�����s
            sql = sql & "CISITN,"   '//����x�X
            sql = sql & "CIKZSB,"   '//�������
            sql = sql & "CIKZNO,"   '//�����ԍ�
            sql = sql & "CIYBTK,"   '//�ʒ��L��
            sql = sql & "CIYBTN,"   '//�ʒ��ԍ�
            sql = sql & "CIKZNM,"   '//�������`�l_�J�i
            sql = sql & "CIERROR,"
            sql = sql & "CIERSR,"
            sql = sql & "CIMCDT,"   '//������   2006/03/24 ADD
            sql = sql & "CIUSID,"   '//�X�V��
            sql = sql & "CIUPDT,"   '//�X�V��
            sql = sql & "CIOKFG " & vbCrLf  '//�捞�n�j�t���O
            sql = sql & ")VALUES(" & vbCrLf
            sql = sql & "TO_DATE(" & gdDBS.ColumnDataSet(insDate, "D", vEnd:=True) & ",'yyyy-mm-dd hh24:mi:ss'),"
            sql = sql & "sqImportSeq.NEXTVAL,"
            sql = sql & " (SELECT ABITKB FROM taItakushaMaster WHERE ABKYTP = '" & Left(Hogosha.Keiyakusha, 1) & "'),"
            sql = sql & gdDBS.ColumnDataSet(Hogosha.Keiyakusha)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.Kyoshittsu)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.HogoshaNo)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.HogoshaKanji)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.HogoshaKana)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.SeitoShimei)
'//2006/04/26 ���z�Ȃ̂� NULL �ł͖��� �u�O�v��������
            sql = sql & gdDBS.ColumnDataSet(gdDBS.Nz(Hogosha.FurikaeGaku, 0), "L") & vbCrLf
            sql = sql & gdDBS.ColumnDataSet(Hogosha.BankName)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.ShitenName)
            If "" <> Trim(Hogosha.BankCode) _
            And "" <> Trim(Hogosha.ShitenCode) Then     '//���ԋ��Z�@�փR�[�h �L������
                sql = sql & gdDBS.ColumnDataSet(eBankKubun.KinnyuuKikan, "I")   '//���ԋ��Z�@��
            ElseIf "" <> Trim(Hogosha.TuuchoKigou) _
                And "" <> Trim(Hogosha.TuuchoBango) Then '//�X�֋Ǐ�� �L������
                sql = sql & gdDBS.ColumnDataSet(eBankKubun.YuubinKyoku, "I")   '//�X�֋�
            Else
                sql = sql & "NULL,"   '//���Z�@�֋敪��NULL
            End If
            sql = sql & gdDBS.ColumnDataSet(Hogosha.BankCode)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.ShitenCode)
            sql = sql & gdDBS.ColumnDataSet(Val(Hogosha.YokinShumoku))  '//�a����ځ��O�̑Ή�
            sql = sql & gdDBS.ColumnDataSet(Hogosha.KouzaBango)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.TuuchoKigou)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.TuuchoBango)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.HogoshaKana)
            sql = sql & gdDBS.ColumnDataSet(mRimp.errImport) & vbCrLf
            sql = sql & gdDBS.ColumnDataSet(mRimp.errImport) & vbCrLf
            sql = sql & gdDBS.ColumnDataSet(Hogosha.MochikomiBi, "L") & vbCrLf    '//������
            sql = sql & gdDBS.ColumnDataSet(gdDBS.LoginUserName)
            sql = sql & "SYSDATE,"
            sql = sql & gdDBS.ColumnDataSet(mRimp.updInvalid, "I", vEnd:=True)
            sql = sql & ")"
            Call gdDBS.Database.ExecuteSQL(sql)
#If DATA_DUPLICATE = True Then  '//�U�ֈ˗����͏d�����`�F�b�N
''''''''''''        End If
''''''''''''        Call dyn.Close
''''''''''''        Set dyn = Nothing
#End If
    Loop
    '//�捞���ʂ̍ŏI�ҏW
    '//2006/04/26 �ی�Ҕԍ��A�����ԍ��A�ʒ��L���A�ʒ��ԍ��̑O�[����Ԓǉ�
    sql = "UPDATE " & mRimp.TcHogoshaImport & " a SET "
    sql = sql & "CIKSCD = DECODE(CIKSCD,NULL,NULL,LPAD(CIKSCD,3,'0'))," & vbCrLf    '//�����ԍ��F       ���͂��P���Ȃ̂œ��͂��L��ꍇ�݂̂R���ɕҏW
    sql = sql & "CIHGCD = DECODE(CIHGCD,NULL,NULL,LPAD(CIHGCD,4,'0'))," & vbCrLf    '//�ی�ҁF
    sql = sql & "CIBANK = DECODE(CIBANK,NULL,NULL,LPAD(CIBANK,4,'0'))," & vbCrLf    '//���Z�@�փR�[�h�F ���͂��S���������͂��L��ꍇ�݂̂S���ɕҏW
    sql = sql & "CISITN = DECODE(CISITN,NULL,NULL,LPAD(CISITN,3,'0'))," & vbCrLf    '//�x�X�R�[�h�F     ���͂��R���������͂��L��ꍇ�݂̂R���ɕҏW
    sql = sql & "CIKZNO = DECODE(CIKZNO,NULL,NULL,LPAD(CIKZNO,7,'0'))," & vbCrLf    '//�����ԍ� �V��
    sql = sql & "CIYBTK = DECODE(CIYBTK,NULL,NULL,LPAD(CIYBTK," & mRimp.YubinKigouLength & ",'0'))," & vbCrLf     '//�ʒ��L�� �R��
    sql = sql & "CIYBTN = DECODE(CIYBTN,NULL,NULL,LPAD(CIYBTN," & mRimp.YubinBangoLength & ",'0')) " & vbCrLf     '//�ʒ��ԍ� �W��
    sql = sql & " WHERE CIINDT = TO_DATE(" & gdDBS.ColumnDataSet(insDate, "D", vEnd:=True) & ",'yyyy-mm-dd hh24:mi:ss')"
    Call gdDBS.Database.ExecuteSQL(sql)
    Close #fp
    '//�X�e�[�^�X�s�̐���E����
    stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�捞����(" & recCnt & "��)"
    pgrProgressBar.Value = pgrProgressBar.Max
    '//�U���˗����f�[�^�̈ʒu�����W�X�g���ɕۊ�
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
    If err Then
        Dim errCode As Integer, errMsg As String
        If gdDBS.Database.LastServerErr Then
            errCode = gdDBS.Database.LastServerErr
            errMsg = gdDBS.Database.LastServerErrText
        Else
            errCode = err
            errMsg = Error
        End If
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�捞�G���[(" & errCode & ")"
        Call gdDBS.AutoLogOut(mCaption, recCnt & "���ڂŃG���[�������������ߎ捞�����͒��~����܂����B(Error=" & errMsg & ")")
    Else
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�捞���f"
        Call gdDBS.AutoLogOut(mCaption, "�捞�����͒��~����܂����B")
    End If
    Call pLockedControl(True)
    GoTo cmdImport_ClickAbort:
End Sub

'//���Z�@�ց��x�X���̃}�b�`���O�p
Private Function pCompare(vElm1 As Variant, vElm2 As Variant, Optional vCutString As Variant = "") As Boolean
    '// vElm1 �� vElm2 �������ł���� True
    '//Replace()�ȊO�ł��悤�Ƃ���Ƃ�₱�����̂ŁI�I�I�~�߁B
    pCompare = Replace(vElm1, vCutString, "") = Replace(vElm2, vCutString, "")
End Function

Private Function pErrorCount() As Integer
    On Error GoTo pErrorCountError
    pErrorCount = UBound(mErrStts)
    Exit Function
pErrorCountError:
    pErrorCount = -1
End Function

Private Sub pSetErrorStatus(vField As Variant, vError As Integer, Optional vMsg As String = "")
    On Error GoTo SetErrorStatusError:
    Dim ix As Integer
    For ix = LBound(mErrStts) To UBound(mErrStts)
        If UCase(vField) = UCase(mErrStts(ix).Field) Then
            If vError < mErrStts(ix).Error Then
                GoTo SetErrorStatusSet:
            End If
            Exit Sub
        End If
    Next ix
    ix = UBound(mErrStts) + 1
    ReDim Preserve mErrStts(ix) As tpErrorStatus
SetErrorStatusSet:
    mErrStts(ix).Field = UCase(vField)
    mErrStts(ix).Error = vError
    If "" <> vMsg Then
        mErrStts(ix).Message = vMsg
    End If
    Exit Sub
SetErrorStatusError:
    ix = 0
    ReDim Preserve mErrStts(0 To 0) As tpErrorStatus
    GoTo SetErrorStatusSet:
End Sub

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

'/////////////////////////////////////////////////////////////////////////
'//�ʂɂP��������
Public Function gDataCheck(vImpDate As Variant, Optional vSeqNo As Long = -1) As Boolean
    Dim Block As Integer, sqlStep As Long
    Const cMaxBlock As Integer = 5
    Block = cMaxBlock
#If BLOCK_CHECK = True Then           '//�`�F�b�N���̃u���b�N���������邩�H��\���F�f�o�b�N���̂�
    mCheckBlocks = 0
#End If
    
    '// WHERE ��ɂ͕K���t��
    Dim SameConditions As String
    SameConditions = " AND CIINDT = TO_DATE('" & vImpDate & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf
    'SameConditions = " AND CIOKFG NOT IN(" & mRimp.updInvalid & "," & mRimp.updWarnErr & ")" & vbCrLf
    'SameConditions = " AND CIERROR = " & mRimp.errNormal
    If -1 <> vSeqNo Then
        SameConditions = SameConditions & vbCrLf & " AND CISEQN = " & vSeqNo
    End If
    
    On Error GoTo gDataCheckError:
    
    Dim ms As New MouseClass
    Call ms.Start
    fraProgressBar.Visible = True
    
    Call gdDBS.AutoLogOut(mCaption, "[" & vImpDate & ":" & vSeqNo & "] �̃`�F�b�N�������J�n����܂����B")
    
    Call gdDBS.Database.BeginTrans          '//�g�����U�N�V�����J�n

    '////////////////////////////////////////
    '//�폜���ă`�F�b�N���镶�����`
    Dim BankCutName As Variant, ShitenCutName As Variant
    Dim updFlag As Integer, impName As String, mstName As String
    '//��s����
    BankCutName = Array("", "��s", "�M�p����", "�M�p�g��", _
                            "�J������", "�����g��", "�_�Ƌ����g��", _
                            "���Ƌ����g���A����")
    '//�x�X����
    ShitenCutName = Array("", "�x�X", "�o����", "�c�ƕ�", "�x��")
    Dim sql As String, recCnt As Long, sysDate As String
    Dim ix As Integer, msg As String
#If ORA_DEBUG = 1 Then
    Dim dynM As OraDynaset, dynS As OraDynaset
#Else
    Dim dynM As Object, dynS As Object
#End If
    sysDate = gdDBS.sysDate("YYYYMMDD")
    '//////////////////////////////////////////////////
    '//�G���[���ڃ��Z�b�g
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mRimp.TcHogoshaImport & " a SET " & vbCrLf
    sql = sql & mRimp.StatusColumns(" = " & mRimp.errNormal & "," & vbCrLf)
    '//���ƂŌx���f�[�^���u�}�X�^���f����v�Ƃ��Ă���f�[�^������̂ŏ��������Ȃ�
    '//2006/03/14 ��C���������͂��̂܂܂ɂ��āu�O�v�ɒu����
    sql = sql & " CIOKFG = CASE WHEN CIOKFG >= " & mRimp.updWarnUpd & " THEN CIOKFG" & vbCrLf
    sql = sql & "               ELSE " & mRimp.updNormal & vbCrLf
    sql = sql & "          END,"
    sql = sql & " CIWMSG = NULL,"   '//���[�j���O���b�Z�[�W
    sql = sql & " CIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " CIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE 1 = 1" & vbCrLf '//���܂��Ȃ�
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '////////////////////////////////////////////
    '//�U�ֈ˗������P������������
    sql = "SELECT a.* " & vbCrLf
    sql = sql & " FROM " & mRimp.TcHogoshaImport & " a " & vbCrLf
    sql = sql & " WHERE 1 = 1" & vbCrLf
    sql = sql & SameConditions & vbCrLf
    sql = sql & " ORDER BY CIKYCD,CIHGCD,CIKSCD"
    Set dynM = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
    If Not dynM.EOF Then
        pgrProgressBar.Max = dynM.RecordCount
    End If
    Do Until dynM.EOF
        '//////////////////////////////////////////////////
        '// DoEvents �� pProgressBarSet() �̒��Ŏ��s����Ă���
        If False = pProgressBarSet(Block, dynM.RowPosition - 1) Then
            GoTo gDataCheckError:
        End If
        '//���ʂ�������
        Erase mErrStts
        '//////////////////////////////////////////
        '//�ϑ��҃R�[�h�`�F�b�N:�擪�P���� �Q�������A�V�������G���A�C�G��
        sql = "SELECT ABITKB " & vbCrLf
        sql = sql & " FROM taItakushaMaster   a " & vbCrLf
        sql = sql & " WHERE ABKYTP = " & gdDBS.ColumnDataSet(Left(dynM.Fields("CIKYCD"), 1), vEnd:=True) & vbCrLf
        Set dynS = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
        If dynS.EOF Then
            Call pSetErrorStatus("CIITKBE", mRimp.errInvalid, "�ϑ��҂��Ԉ���Ă��܂�.")
        End If
        Call dynS.Close
        Set dynS = Nothing
        '//////////////////////////////////////////
        '//�_��҃R�[�h�`�F�b�N
        sql = "SELECT BAKYED,BAKYFG " & vbCrLf
        sql = sql & " FROM tbKeiyakushaMaster a " & vbCrLf
        sql = sql & " WHERE (BAITKB,BAKYCD,BASQNO) IN(" & vbCrLf
        sql = sql & "       SELECT BAITKB,BAKYCD,MAX(BASQNO) " & vbCrLf
        sql = sql & "       FROM tbKeiyakushaMaster a" & vbCrLf
        sql = sql & "       WHERE BAITKB = " & gdDBS.ColumnDataSet(dynM.Fields("CIITKB"), vEnd:=True) & vbCrLf
        sql = sql & "         AND BAKYCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIKYCD"), vEnd:=True) & vbCrLf
        sql = sql & "       GROUP BY BAITKB,BAKYCD" & vbCrLf
        sql = sql & "     )" & vbCrLf
        Set dynS = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
        If dynS.EOF Then
            Call pSetErrorStatus("CIKYCDE", mRimp.errInvalid, "�_��҂����݂��܂���.")
        ElseIf dynS.Fields("BAKYED") < sysDate Or 0 <> Val(gdDBS.Nz(dynS.Fields("BAKYFG"))) Then
            Call pSetErrorStatus("CIKYCDE", mRimp.errInvalid, "�_��҂�����Ԃł�.")
        End If
        Call dynS.Close
        Set dynS = Nothing
        '//################################################
        '//2008/10/14 ���r�u�����N���m
        If 0 <> InStr(dynM.Fields("CIHGCD"), " ") Then
            Call pSetErrorStatus("CIHGCDE", mRimp.errInvalid, "�ی�Ҕԍ��Ƀu�����N������܂�.")
        End If
        If dynM.Fields("CIKKBN") = eBankKubun.KinnyuuKikan Then
            If 0 <> InStr(dynM.Fields("CIBANK"), " ") Then
                Call pSetErrorStatus("CIBANKE", mRimp.errInvalid, "���Z�@�ւɃu�����N������܂�.")
            End If
            If 0 <> InStr(dynM.Fields("CISITN"), " ") Then
                Call pSetErrorStatus("CISITNE", mRimp.errInvalid, "�x�X�Ƀu�����N������܂�.")
            End If
            If 0 <> InStr(dynM.Fields("CIKZNO"), " ") Then
                Call pSetErrorStatus("CIKZNOE", mRimp.errInvalid, "�����ԍ��Ƀu�����N������܂�.")
            End If
        ElseIf dynM.Fields("CIKKBN") = eBankKubun.YuubinKyoku Then
            If 0 <> InStr(dynM.Fields("CIYBTK"), " ") Then
                Call pSetErrorStatus("CIYBTKE", mRimp.errInvalid, "�ʒ��L���Ƀu�����N������܂�.")
            End If
            If 0 <> InStr(dynM.Fields("CIYBTN"), " ") Then
                Call pSetErrorStatus("CIYBTNE", mRimp.errInvalid, "�ʒ��ԍ��Ƀu�����N������܂�.")
            End If
        End If
        '//2008/10/14 ���r�u�����N���m
        '//################################################
        '//�����ԍ��`�F�b�N
        If IsNull(dynM.Fields("CIKSCD")) Then
            Call pSetErrorStatus("CIKSCDE", mRimp.errInvalid, "�����ԍ��������͂ł�.")
        End If
        '//////////////////////////////////////////
        '//�ی�Ҕԍ��`�F�b�N
        If IsNull(dynM.Fields("CIHGCD")) Then
            Call pSetErrorStatus("CIHGCDE", mRimp.errInvalid, "�ی�Ҕԍ��������͂ł�.")
        End If
        '//////////////////////////////////////////
        '//�ی�Җ�(����)�`�F�b�N
        If IsNull(dynM.Fields("CIKJNM")) Then
            Call pSetErrorStatus("CIKJNME", mRimp.errInvalid, "�ی�Җ�(����)�������͂ł�.")
        End If
        '//////////////////////////////////////////
        '//�ی�Җ�(�J�i)�`�F�b�N
        If IsNull(dynM.Fields("CIKNNM")) Then
            Call pSetErrorStatus("CIKNNME", mRimp.errInvalid, "�ی�Җ�(�J�i)�������͂ł�.")
        End If
        '//////////////////////////////////////////
        '//�ߋ�/���� �U�ֈ˗����E�捞�f�[�^�Ƃ̃`�F�b�N
        sql = "SELECT MAX(DupCode) DUPCODE FROM(" & vbCrLf
        sql = sql & " SELECT " & gdDBS.ColumnDataSet("�ߋ�", vEnd:=True) & " DupCode " & vbCrLf
        sql = sql & " FROM " & mRimp.TcHogoshaImport & " a " & vbCrLf
        sql = sql & " WHERE CIINDT <>TO_DATE('" & vImpDate & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf
        sql = sql & "   AND CIITKB = " & gdDBS.ColumnDataSet(dynM.Fields("CIITKB"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CIKYCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIKYCD"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CIKSCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIKSCD"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CIHGCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIHGCD"), vEnd:=True) & vbCrLf
        sql = sql & " UNION " & vbCrLf
        sql = sql & " SELECT " & gdDBS.ColumnDataSet("����", vEnd:=True) & " DupCode " & vbCrLf
        sql = sql & " FROM " & mRimp.TcHogoshaImport & " a " & vbCrLf
        sql = sql & " WHERE CIINDT = TO_DATE('" & vImpDate & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf
        '//�������g�ȊO
        sql = sql & "   AND CISEQN <>" & gdDBS.ColumnDataSet(dynM.Fields("CISEQN"), "I", vEnd:=True) & vbCrLf
        sql = sql & "   AND CIITKB = " & gdDBS.ColumnDataSet(dynM.Fields("CIITKB"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CIKYCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIKYCD"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CIKSCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIKSCD"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CIHGCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIHGCD"), vEnd:=True) & vbCrLf
        sql = sql & ")"
        Set dynS = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
        '//MAX() �ł��Ă���̂ŕK�����݂���
        If Not IsNull(dynS.Fields("DupCode")) Then
            Call pSetErrorStatus("CIHGCDE", mRimp.errWarning, dynS.Fields("DupCode") & "�̎捞�f�[�^�ɑ��݂��܂�.")
        End If
        Call dynS.Close
        Set dynS = Nothing
        '//////////////////////////////////////////
        '//�ی�҃}�X�^�Ƃ̃`�F�b�N
        sql = "SELECT a.* " & vbCrLf
        sql = sql & " FROM tcHogoshaMaster a " & vbCrLf
        sql = sql & " WHERE (CAITKB,CAKYCD,CAKSCD,CAHGCD,CASQNO) IN(" & vbCrLf
        sql = sql & "       SELECT CAITKB,CAKYCD,CAKSCD,CAHGCD,MAX(CASQNO) " & vbCrLf
        sql = sql & "       FROM tcHogoshaMaster a" & vbCrLf
        sql = sql & "       WHERE CAITKB = " & gdDBS.ColumnDataSet(dynM.Fields("CIITKB"), vEnd:=True) & vbCrLf
        sql = sql & "         AND CAKYCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIKYCD"), vEnd:=True) & vbCrLf
        sql = sql & "         AND CAKSCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIKSCD"), vEnd:=True) & vbCrLf
        sql = sql & "         AND CAHGCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIHGCD"), vEnd:=True) & vbCrLf
        sql = sql & "       GROUP BY CAITKB,CAKYCD,CAKSCD,CAHGCD" & vbCrLf
        sql = sql & "     )" & vbCrLf
        Set dynS = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
        '//////////////////////////////////////////
        '//�f�[�^������ꍇ�݂̂ŉF�������͖����F�r�s�`�q�s
        If Not dynS.EOF Then
            If dynS.Fields("CAKYED") < sysDate Or 0 <> Val(gdDBS.Nz(dynS.Fields("CAKYFG"))) Then
                Call pSetErrorStatus("CIHGCDE", mRimp.errWarning, "�ی�҃}�X�^�͉���Ԃł�.")
            Else
                Call pSetErrorStatus("CIHGCDE", mRimp.errWarning, "�ی�҃}�X�^�Ɋ��ɑ��݂��܂�.")
            End If
            '//////////////////////////////////////////
            '//�ی�Җ�(����)�`�F�b�N
            If Replace(Replace(dynM.Fields("CIKJNM"), "�@", ""), " ", "") _
            <> Replace(Replace(dynS.Fields("CAKJNM"), "�@", ""), " ", "") Then
                Call pSetErrorStatus("CIKJNME", mRimp.errWarning, "�ی�Җ�(����)�ɑ��Ⴊ����܂�.")
            End If
            '//////////////////////////////////////////
            '//�ی�Җ�(�J�i)�`�F�b�N
            '//2007/04/20 �p���`�ɕی�҃J�i NULL �L��̈׃G���[
            If Not IsNull(dynM.Fields("CIKNNM")) Then
                If Replace(Replace(dynM.Fields("CIKNNM"), "�@", ""), " ", "") _
                <> Replace(Replace(dynS.Fields("CAKNNM"), "�@", ""), " ", "") Then
                    Call pSetErrorStatus("CIKNNME", mRimp.errWarning, "�ی�Җ�(�J�i)�ɑ��Ⴊ����܂�.")
                End If
            End If
            If dynM.Fields("CIKKBN") = eBankKubun.KinnyuuKikan Then
                '//////////////////////////////////////////
                '//���Z�@�փ`�F�b�N
                If dynM.Fields("CIBANK") <> dynS.Fields("CABANK") Then
                    Call pSetErrorStatus("CIBANKE", mRimp.errWarning, "���Z�@�ւɑ��Ⴊ����܂�.")
                End If
                '//////////////////////////////////////////
                '//�x�X�`�F�b�N
                If dynM.Fields("CISITN") <> dynS.Fields("CASITN") Then
                    Call pSetErrorStatus("CISITNE", mRimp.errWarning, "�x�X�ɑ��Ⴊ����܂�.")
                End If
                '//////////////////////////////////////////
                '//�a����ڃ`�F�b�N
                If dynM.Fields("CIKZSB") <> dynS.Fields("CAKZSB") Then
                    Call pSetErrorStatus("CIKZSBE", mRimp.errWarning, "�a����ڂɑ��Ⴊ����܂�.")
                End If
                '//////////////////////////////////////////
                '//�����ԍ��`�F�b�N
                If dynM.Fields("CIKZNO") <> dynS.Fields("CAKZNO") Then
                    Call pSetErrorStatus("CIKZNOE", mRimp.errWarning, "�����ԍ��ɑ��Ⴊ����܂�.")
                End If
            ElseIf dynM.Fields("CIKKBN") = eBankKubun.YuubinKyoku Then
                '//////////////////////////////////////////
                '//�ʒ��L���`�F�b�N
                If dynM.Fields("CIYBTK") <> dynS.Fields("CAYBTK") Then
                    Call pSetErrorStatus("CIYBTKE", mRimp.errWarning, "�ʒ��L���ɑ��Ⴊ����܂�.")
                End If
                '//////////////////////////////////////////
                '//�ʒ��ԍ��`�F�b�N
                If dynM.Fields("CIYBTN") <> dynS.Fields("CAYBTN") Then
                    Call pSetErrorStatus("CIYBTNE", mRimp.errWarning, "�ʒ��ԍ��ɑ��Ⴊ����܂�.")
                End If
            Else
                Call pSetErrorStatus("CIKKBNE", mRimp.errWarning, "���Z�@�֋敪���Ԉ���Ă��܂�.")
            End If
            '//////////////////////////////////////////
            '//�������`�l���`�F�b�N
            If dynM.Fields("CIKZNM") <> dynS.Fields("CAKZNM") Then
                Call pSetErrorStatus("CIKZNME", mRimp.errWarning, "�������`�l���ɑ��Ⴊ����܂�.")
            End If
        End If
        '//�f�[�^������ꍇ�݂̂ŉF�d�m�c
        '//////////////////////////////////////////
        Call dynS.Close
        Set dynS = Nothing
        '//////////////////////////////////////
        '//���Z�@�փ`�F�b�N
        If dynM.Fields("CIKKBN") = eBankKubun.KinnyuuKikan Then
            If IsNull(dynM.Fields("CIBANK")) Then
                Call pSetErrorStatus("CIBANKE", mRimp.errWarning, "���Z�@�փR�[�h�������͂ł�.")
            Else
                sql = "SELECT * FROM tdBankMaster " & vbCrLf
                sql = sql & " WHERE DARKBN = " & gdDBS.ColumnDataSet(eBankRecordKubun.Bank, vEnd:=True) & vbCrLf
                sql = sql & "   AND DABANK = " & gdDBS.ColumnDataSet(dynM.Fields("CIBANK"), vEnd:=True) & vbCrLf
                sql = sql & "   AND DASITN = '000'" & vbCrLf
                sql = sql & " ORDER BY DECODE(DASQNO,':',0,'#',1,'@',2,'''',3,'=',4,9)" & vbCrLf
                Set dynS = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
                If dynS.EOF Then
                    Call pSetErrorStatus("CIBANKE", mRimp.errWarning, "���Z�@�ւ����݂��܂���.")
                Else
                    '//����A�擾�̓��X�|���X���������낤�I����ϐ��ɑ�����ă`�F�b�N
                    impName = gdDBS.Nz(dynM.Fields("CIBKNM"))
                    updFlag = mRimp.errNormal
                    Do Until dynS.EOF
                        mstName = dynS.Fields("DAKJNM")
                        For ix = LBound(BankCutName) To UBound(BankCutName)
                            If True = pCompare(impName, mstName, BankCutName(ix)) Then '//�u�H�H�H�H�v������ă`�F�b�N
                                updFlag = mRimp.errNormal
                                Exit Do    '//�`�F�b�N�n�j
                            Else
                                updFlag = mRimp.errWarning
                            End If
                        Next ix
                        Call dynS.MoveNext
                    Loop
                    If updFlag <> mRimp.errNormal Then
                        Call pSetErrorStatus("CIBKNME", mRimp.errWarning, "���Z�@�֖��̂����v���܂���.")
                        Call pSetErrorStatus("CIBANKE", mRimp.errWarning)
                    End If
                End If
                Call dynS.Close
                Set dynS = Nothing
            End If
            '//////////////////////////////////////
            '//�x�X�`�F�b�N
            If IsNull(dynM.Fields("CISITN")) Then
                Call pSetErrorStatus("CISITNE", mRimp.errWarning, "�x�X�R�[�h�������͂ł�.")
'//2006/07/25 �x�X���`�F�b�N�ɍs���ĂȂ��H�̂� Not �t�^
'//2007/05/23 �x�X���̂̃`�F�b�N�̃f�o�b�O
            ElseIf Not IsNull(dynM.Fields("CIBANK")) Then
                sql = "SELECT * FROM tdBankMaster"
                sql = sql & " WHERE DARKBN = " & gdDBS.ColumnDataSet(eBankRecordKubun.Shiten, vEnd:=True) & vbCrLf
                sql = sql & "   AND DABANK = " & gdDBS.ColumnDataSet(dynM.Fields("CIBANK"), vEnd:=True)
                sql = sql & "   AND DASITN = " & gdDBS.ColumnDataSet(dynM.Fields("CISITN"), vEnd:=True)
                sql = sql & " ORDER BY DASQNO"
                Set dynS = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
                If dynS.EOF Then
                    Call pSetErrorStatus("CISITNE", mRimp.errWarning, "�x�X�����݂��܂���.")
                Else
                    '//����A�擾�̓��X�|���X���������낤�I����ϐ��ɑ�����ă`�F�b�N
                    impName = gdDBS.Nz(dynM.Fields("CISINM"))
                    updFlag = mRimp.errNormal
                    Do Until dynS.EOF
                        mstName = dynS.Fields("DAKJNM")
                        For ix = LBound(ShitenCutName) To UBound(ShitenCutName)
                            If True = pCompare(impName, mstName, ShitenCutName(ix)) Then '//�u�H�H�H�H�v������ă`�F�b�N
                                updFlag = mRimp.errNormal
                                Exit Do    '//�`�F�b�N�n�j
                            Else
                                updFlag = mRimp.errWarning
                            End If
                        Next ix
                        Call dynS.MoveNext
                    Loop
                    If updFlag <> mRimp.errNormal Then
                        Call pSetErrorStatus("CISINME", mRimp.errWarning, "�x�X���̂����v���܂���.")
                        Call pSetErrorStatus("CISITNE", mRimp.errWarning)
                    End If
                End If
                Call dynS.Close
                Set dynS = Nothing
            End If
            '//////////////////////////////////////////
            '//�a����ڃ`�F�b�N
            If dynM.Fields("CIKZSB") = eBankYokinShubetsu.Futsuu _
            Or dynM.Fields("CIKZSB") = eBankYokinShubetsu.Touza Then
            Else
                Call pSetErrorStatus("CIKZSBE", mRimp.errWarning, "�a����ڂɌ�肪����܂�.")
            End If
            '//////////////////////////////////////////
            '//�����ԍ��`�F�b�N
            If "" = gdDBS.Nz(dynM.Fields("CIKZNO")) Then
                Call pSetErrorStatus("CIKZNOE", mRimp.errWarning, "�����ԍ��Ɍ�肪����܂�.")
            End If
        ElseIf dynM.Fields("CIKKBN") = eBankKubun.YuubinKyoku Then
            '//////////////////////////////////////////
            '//�ʒ��L���`�F�b�N
            If IsNull(dynM.Fields("CIYBTK")) Or Len(dynM.Fields("CIYBTK")) < mRimp.YubinKigouLength Then
                Call pSetErrorStatus("CIYBTKE", mRimp.errWarning, "�ʒ��L���Ɍ�肪����܂�.")
            End If
            '//////////////////////////////////////////
            '//�ʒ��ԍ��`�F�b�N
            If IsNull(dynM.Fields("CIYBTN")) Or Len(dynM.Fields("CIYBTN")) < mRimp.YubinBangoLength Then
                Call pSetErrorStatus("CIYBTNE", mRimp.errWarning, "�ʒ��ԍ��Ɍ�肪����܂�.")
            ElseIf "1" <> Right(dynM.Fields("CIYBTN"), 1) Then
                Call pSetErrorStatus("CIYBTNE", mRimp.errWarning, "�ʒ��ԍ��Ɍ�肪����܂�(�������P�ȊO).")
            End If
        Else
            Call pSetErrorStatus("CIKKBNE", mRimp.errWarning, "���Z�@�֋敪�Ɍ�肪����܂�.")
        End If
'//2006/04/26 ���Z�@�ցE�X�֋ǂ̗������͂�����
'//2007/06/12 ���������Ă����͂�����ł���΁H�ǂ����낤�B�Ǝv�������H�H�H
        If "" <> gdDBS.Nz(dynM.Fields("CIYBTK")) & gdDBS.Nz(dynM.Fields("CIYBTK")) _
        And "" <> gdDBS.Nz(dynM.Fields("CIBANK")) & gdDBS.Nz(dynM.Fields("CISITN")) & gdDBS.Nz(dynM.Fields("CIKZNO")) Then
            Call pSetErrorStatus("CIKKBNE", mRimp.errWarning, "���Z�@��/�X�֋ǂ̗����ɓ��͂�����܂�.")
        End If
        
        '////////////////////////////////////////////////
        '//�G���[�̔z�񂪑��݂���� UPDATE ���𐶐�
        If 0 <= pErrorCount() Then
            sql = "UPDATE " & mRimp.TcHogoshaImport & " SET " & vbCrLf
            msg = ""
            For ix = LBound(mErrStts) To UBound(mErrStts)
                msg = msg & mErrStts(ix).Message & vbCrLf
                sql = sql & mErrStts(ix).Field & " = " & mErrStts(ix).Error & "," & vbCrLf
            Next ix
            sql = sql & " CIWMSG = '" & msg & "'," & vbCrLf
            sql = sql & " CIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
            sql = sql & " CIUPDT = SYSDATE" & vbCrLf
            sql = sql & " WHERE CISEQN = " & dynM.Fields("CISEQN") & vbCrLf
            sql = sql & SameConditions & vbCrLf
            recCnt = gdDBS.Database.ExecuteSQL(sql)
        End If
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
    sql = "UPDATE " & mRimp.TcHogoshaImport & " a SET " & vbCrLf
    sql = sql & " CIOKFG =  " & mRimp.updInvalid & "," & vbCrLf    '//�}�X�^���f�s��
    sql = sql & " CIERROR = " & mRimp.errInvalid & "," & vbCrLf
    sql = sql & " CIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " CIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE(" & vbCrLf
    sql = sql & mRimp.StatusColumns(" = " & mRimp.errInvalid & vbCrLf & " OR ", Len(vbCrLf & " OR ")) & vbCrLf & ")" & vbCrLf
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//�S�̃G���[���ڃZ�b�g�F�ŏ��ɐ���ɂ��Ă���̂Łu����v�t���O�͕s�v
    '//�x���f�[�^�F�}�X�^���f���Ȃ��f�[�^
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mRimp.TcHogoshaImport & " a SET " & vbCrLf
    sql = sql & " CIOKFG =  " & mRimp.updWarnErr & "," & vbCrLf   '//�}�X�^���f���Ȃ��t���O
    sql = sql & " CIERROR = " & mRimp.errWarning & "," & vbCrLf
    sql = sql & " CIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " CIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE CIERROR = " & mRimp.errNormal & vbCrLf    '//�ُ�Ŗ���
    sql = sql & "   AND CIOKFG <= " & mRimp.updNormal & vbCrLf
    sql = sql & "   AND(" & vbCrLf
    sql = sql & mRimp.StatusColumns(" >= " & mRimp.errWarning & vbCrLf & " OR ", Len(vbCrLf & " OR ")) & vbCrLf & ")" & vbCrLf
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//�\�[�g�p�� CIERROR=>CIERSR �ɃR�s�[
    '//Spread�ŉ��z���[�h�ɂ���ƃ��A���ɕς��ׁA�C����=CIEROR�A�Œ蕔=CIERSR �Ƃ���
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mRimp.TcHogoshaImport & " a SET " & vbCrLf
    sql = sql & " CIERSR = CIERROR "
    sql = sql & " WHERE 1 = 1"  '//���܂��Ȃ�
    sql = sql & SameConditions & vbCrLf
    If -1 <> vSeqNo Then        '//�s�w�莞�ɂ͍X�V���Ȃ����܂��Ȃ��F���z���[�h�Ń��A���ɕς���
        sql = sql & " AND 1 = -1"
    End If
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    
    Call gdDBS.Database.CommitTrans         '//�g�����U�N�V��������I��
    fraProgressBar.Visible = False
    Call gdDBS.AutoLogOut(mCaption, "[" & vImpDate & ":" & vSeqNo & "] �̃`�F�b�N�������������܂����B")
    '//�X�e�[�^�X�s�̐���E����
    stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�`�F�b�N����"
    gDataCheck = True
    
#If BLOCK_CHECK = True Then           '//�`�F�b�N���̃u���b�N���������邩�H��\���F�f�o�b�N���̂�
     Call MsgBox("�`�F�b�N�����u���b�N�� " & mCheckBlocks & " �ӏ��ł����B")
#End If
    
    Exit Function
gDataCheckError:
    fraProgressBar.Visible = False
    Call gdDBS.Database.Rollback            '//�g�����U�N�V�����ُ�I��
    If err Then
        Dim errCode As Integer, errMsg As String
        If gdDBS.Database.LastServerErr Then
            errCode = gdDBS.Database.LastServerErr
            errMsg = gdDBS.Database.LastServerErrText
        Else
            errCode = err
            errMsg = Error
        End If
        fraProgressBar.Visible = False
        '//�X�e�[�^�X�s�̐���E����
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�`�F�b�N�G���[(" & errCode & ")"
        Call gdDBS.AutoLogOut(mCaption, "[" & vImpDate & ":" & vSeqNo & "] �̃`�F�b�N�������ɃG���[���������܂����B(Error=" & errCode & ")")
        Call MsgBox("�`�F�b�N�Ώ� = [" & cboImpDate.Text & "]" & vbCrLf & _
                    "�̓G���[�������������߃`�F�b�N�͒��~����܂����B" & vbCrLf & errMsg, _
                vbOKOnly + vbCritical, mCaption)
    Else
        Call gdDBS.AutoLogOut(mCaption, "[" & vImpDate & ":" & vSeqNo & "] �̃`�F�b�N���������f����܂����B")
        '//�X�e�[�^�X�s�̐���E����
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "�`�F�b�N���f"
    End If
End Function

Private Sub cmdCheck_Click()
    If False = pCheckSubForm Then
        Exit Sub
    End If
    If -1 <> pAbortButton(cmdCheck, cBtnCheck) Then
        Exit Sub
    End If
    cmdCheck.Caption = cBtnCancel
    '//�R�}���h�E�{�^������
    Call pLockedControl(False, cmdCheck)
    '//�`�F�b�N����
    If True = gDataCheck(cboImpDate.Text) Then
        '//�f�[�^�ǂݍ��݁� Spread �ɐݒ蔽�f
        Call pReadDataAndSetting
    End If
    '//�{�^����߂�
    cmdCheck.Caption = cBtnCheck
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
End Sub

#If ORA_DEBUG = 1 Then
Private Function pHogoshaInsert(vInDyn As OraDynaset) As Boolean
#Else
Private Function pHogoshaInsert(vInDyn As Object) As Boolean
#End If
    Dim sql As String
    sql = "INSERT INTO tcHogoshaMaster ( " & vbCrLf
    sql = sql & "CAITKB," & vbCrLf  '//�ϑ��ҋ敪
    sql = sql & "CAKYCD," & vbCrLf  '//�_��Ҕԍ�
    sql = sql & "CAKSCD," & vbCrLf  '//�����ԍ�
    sql = sql & "CAHGCD," & vbCrLf  '//�ی�Ҕԍ�
    sql = sql & "CASQNO," & vbCrLf  '//�ی�҂r�d�p
    sql = sql & "CAKJNM," & vbCrLf  '//�ی�Җ�_����
    sql = sql & "CAKNNM," & vbCrLf  '//�ی�Җ�_�J�i
    sql = sql & "CASTNM," & vbCrLf  '//���k����
    sql = sql & "CAKKBN," & vbCrLf  '//������Z�@�֋敪
    sql = sql & "CABANK," & vbCrLf  '//�����s
    sql = sql & "CASITN," & vbCrLf  '//����x�X
    sql = sql & "CAKZSB," & vbCrLf  '//�������
    sql = sql & "CAKZNO," & vbCrLf  '//�����ԍ�
    sql = sql & "CAYBTK," & vbCrLf  '//�ʒ��L��
    sql = sql & "CAYBTN," & vbCrLf  '//�ʒ��ԍ�
    sql = sql & "CAKZNM," & vbCrLf  '//�������`�l_�J�i
    sql = sql & "CAKYST," & vbCrLf  '//�_��J�n��
    sql = sql & "CAKYED," & vbCrLf  '//�_��I����
    sql = sql & "CAFKST," & vbCrLf  '//�U�֊J�n��
    sql = sql & "CAFKED," & vbCrLf  '//�U�֏I����
    sql = sql & "CASKGK," & vbCrLf  '//�����\��z
    sql = sql & "CAHKGK," & vbCrLf  '//�ύX����z
    sql = sql & "CAKYDT," & vbCrLf  '//����
    sql = sql & "CAKYFG," & vbCrLf  '//���t���O
    sql = sql & "CATRFG," & vbCrLf  '//�`���X�V�t���O
    sql = sql & "CAUSID," & vbCrLf  '//�쐬��
    sql = sql & "CAADDT," & vbCrLf  '//�X�V��
    sql = sql & "CANWDT " & vbCrLf  '//�V�K�f�[�^������
    sql = sql & ") SELECT " & vbCrLf
    sql = sql & "CiITKB," & vbCrLf  '//�ϑ��ҋ敪
    sql = sql & "CiKYCD," & vbCrLf  '//�_��Ҕԍ�
    sql = sql & "CiKSCD," & vbCrLf  '//�����ԍ�
    sql = sql & "CiHGCD," & vbCrLf  '//�ی�Ҕԍ�
    sql = sql & "TO_CHAR(SYSDATE,'yyyymmdd')," & vbCrLf  '//�ی�҂r�d�p
    sql = sql & "CiKJNM," & vbCrLf  '//�ی�Җ�_����
    sql = sql & "CiKNNM," & vbCrLf  '//�ی�Җ�_�J�i
    sql = sql & "CiSTNM," & vbCrLf  '//���k����
    sql = sql & "CiKKBN," & vbCrLf  '//������Z�@�֋敪
    sql = sql & "CiBANK," & vbCrLf  '//�����s
    sql = sql & "CiSITN," & vbCrLf  '//����x�X
    sql = sql & "CiKZSB," & vbCrLf  '//�������
    sql = sql & "CiKZNO," & vbCrLf  '//�����ԍ�
    sql = sql & "CiYBTK," & vbCrLf  '//�ʒ��L��
    sql = sql & "CiYBTN," & vbCrLf  '//�ʒ��ԍ�
    sql = sql & "CiKZNM," & vbCrLf  '//�������`�l_�J�i
    sql = sql & "     0," & vbCrLf  '//�_��J�n��
    sql = sql & "20991231," & vbCrLf  '//�_��I����
    sql = sql & "     0," & vbCrLf  '//�U�֊J�n��
    sql = sql & "20991231," & vbCrLf  '//�U�֏I����
    sql = sql & "CiSKGK," & vbCrLf  '//�����\��z
    sql = sql & "  NULL," & vbCrLf  '//�ύX����z
    sql = sql & "  NULL," & vbCrLf  '//����
    sql = sql & "     0," & vbCrLf  '//���t���O
    sql = sql & "  NULL," & vbCrLf  '//�`���X�V�t���O
    sql = sql & gdDBS.ColumnDataSet(MainModule.gcImportHogoshaUser) & vbCrLf    '//�X�V�҂h�c
    sql = sql & "SYSDATE," & vbCrLf  '//�X�V��
    sql = sql & "   NULL " & vbCrLf  '//�V�K�f�[�^������
    sql = sql & " FROM " & mRimp.TcHogoshaImport
    sql = sql & " WHERE CIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss') " & vbCrLf
    sql = sql & "   AND CIKYCD = " & gdDBS.ColumnDataSet(vInDyn.Fields("CIKYCD"), vEnd:=True)
    sql = sql & "   AND CIHGCD = " & gdDBS.ColumnDataSet(vInDyn.Fields("CIHGCD"), vEnd:=True)
    sql = sql & "   AND CISEQN = " & gdDBS.ColumnDataSet(vInDyn.Fields("CISEQN"), vEnd:=True)
    Call gdDBS.Database.ExecuteSQL(sql)
    pHogoshaInsert = True
End Function

#If ORA_DEBUG = 1 Then
Private Function pHogoshaUpdate(vOutDyn As OraDynaset, vInDyn As OraDynaset) As Boolean
#Else
Private Function pHogoshaUpdate(vOutDyn As Object, vInDyn As Object) As Boolean
#End If
    Dim Fields As Variant, ix As Integer, chg As Boolean
    Dim sql As String
    Fields = Array("CaITKB", "CaKYCD", "CaKSCD", "CaHGCD", "CaKJNM", "CaKNNM", "CaSTNM", "CaKKBN", _
                   "CaBANK", "CaSITN", "CaKZSB", "CaKZNO", "CaYBTK", "CaYBTN", "CaKZNM", "CaSKGK")
    '//���͂̑��ᕪ�̂ݍX�V����
    For ix = LBound(Fields) To UBound(Fields)
        chg = False
        '//2007/04/24 ���肪 NULL �ł���ƈႤ�Ɣ��f����čX�V���鍀�ڂł͂Ȃ��Ȃ�o�O�C��
        If IsNull(vOutDyn.Fields(Fields(ix))) And Not IsNull(vInDyn.Fields("Ci" & Mid(Fields(ix), 3))) Then
            '//�o�͐悪�Е� NULL
            chg = True
        ElseIf Not IsNull(vOutDyn.Fields(Fields(ix))) And IsNull(vInDyn.Fields("Ci" & Mid(Fields(ix), 3))) Then
            '//���͐悪�Е� NULL
            chg = True
        ElseIf vOutDyn.Fields(Fields(ix)) <> vInDyn.Fields("Ci" & Mid(Fields(ix), 3)) Then
            '//�o�͐�Ɠ��͐�ɑ��Ⴊ�L��
            chg = True
        End If
        If True = chg Then
            sql = sql & Fields(ix) & " = " & gdDBS.ColumnDataSet(vInDyn.Fields("Ci" & Mid(Fields(ix), 3)), "S") & vbCrLf
        End If
    Next ix
'//�p���`�f�[�^�Ƃ̌���������Ȃ��Ȃ�̂ł�߂��F��ɉ��炩�͍X�V����
#If 0 Then
    '//�������łȂ��A���ׂĂ̗�ɕύX��������΍X�V���Ȃ�
    If mRimp.updResetCancel <> vInDyn.Fields("CiOKFG") And "" = sql Then
        pHogoshaUpdate = True
        Exit Function
    End If
#End If
    sql = "UPDATE tcHogoshaMaster SET " & sql   '//��Œ�`�����\�����u�Ō�Ɂv�ɕt��
    If mRimp.updResetCancel = vInDyn.Fields("CiOKFG") Then
        sql = sql & " CAKYED = CASE WHEN CAKYED < 20991231 THEN 20991231 END," & vbCrLf
        sql = sql & " CAFKED = CASE WHEN CAFKED < 20991231 THEN 20991231 END," & vbCrLf
        sql = sql & " CAKYDT = NULL," & vbCrLf
        sql = sql & " CAKYFG = 0," & vbCrLf
    End If
    sql = sql & " CAUSID = " & gdDBS.ColumnDataSet(MainModule.gcImportHogoshaUser) & vbCrLf
    sql = sql & " CAUPDT = SYSDATE" & vbCrLf
    '//���ɍX�V����ׂ��Y�����R�[�h�͓ǂݏo���ς�
    sql = sql & " WHERE CAKYCD = " & gdDBS.ColumnDataSet(vOutDyn.Fields("CAKYCD"), vEnd:=True) & vbCrLf
    sql = sql & "   AND CAKSCD = " & gdDBS.ColumnDataSet(vOutDyn.Fields("CAKSCD"), vEnd:=True) & vbCrLf
    sql = sql & "   AND CAHGCD = " & gdDBS.ColumnDataSet(vOutDyn.Fields("CAHGCD"), vEnd:=True) & vbCrLf
    sql = sql & "   AND CASQNO = " & gdDBS.ColumnDataSet(vOutDyn.Fields("CASQNO"), "L", vEnd:=True) & vbCrLf
    Call gdDBS.Database.ExecuteSQL(sql)
    pHogoshaUpdate = True
End Function

Private Sub cmdUpdate_Click()
    If False = pCheckSubForm Then
        Exit Sub
    End If
    If -1 <> pAbortButton(cmdUpdate, cBtnUpdate) Then
        Exit Sub
    End If
    If vbOK <> MsgBox("�}�X�^�̔��f���J�n���܂��B" & vbCrLf & vbCrLf & "��낵���ł����H", vbOKCancel + vbInformation, Me.Caption) Then
        Exit Sub
    End If
    cmdUpdate.Caption = cBtnCancel
    '//�R�}���h�E�{�^������
    Call pLockedControl(False, cmdUpdate)
    
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset, msg As String
#Else
    Dim sql As String, dyn As Object, msg As String
#End If
    '//////////////////////////////////////////////////////////
    '//�����Ŏg�p���鋤�ʂ� WHERE ����
    Dim Condition As String
    Condition = Condition & " AND CIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss') " & vbCrLf
    '// CIERROR >= 0 AND CIOKFG >= 0 �ł��邱��
    Condition = Condition & " AND CIERROR >= 0" & vbCrLf
    Condition = Condition & " AND CIOKFG  >= 0"
    Condition = Condition & " AND CIMUPD   = 0" '//2006/04/04 �}�X�^���f�n�j�t���O���ڒǉ�
    '///////////////////////////////////////
    '// �捞�����P�ʂ� TcHogoshaImport ���ɓ����ی�҂����݂��Ȃ�����
    '//2006/03/17 �d���f�[�^�͌㏟���ōX�V����悤�ɕύX�ɂ����̂ł��肦�Ȃ����낤�H
    '//2006/04/24 �����ԍ���ǉ�
    sql = " SELECT CIKYCD,CIKSCD,CIHGCD"
    sql = sql & " FROM " & mRimp.TcHogoshaImport
    sql = sql & " WHERE 1 = 1"  '//���܂��Ȃ�
    sql = sql & Condition
    sql = sql & " GROUP BY CIKYCD,CIKSCD,CIHGCD"
    sql = sql & " HAVING COUNT(*) > 1 "     '//����̕ی�҂����݂��邩�H
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If Not dyn.EOF Then
        msg = "�捞���� [ " & cboImpDate.Text & " ] ����" & vbCrLf & _
              "�@ �ی�� [ " & dyn.Fields("CIKYCD") & " - " & dyn.Fields("CIHGCD") & " ] ���������݂����     " & vbCrLf & _
              "�}�X�^���f�͏������s���o���܂���B"
    End If
    Call dyn.Close
    Set dyn = Nothing
    If "" <> msg Then
        Call MsgBox(msg, vbOKOnly + vbCritical, mCaption)
        '//�{�^����߂�
        cmdUpdate.Caption = cBtnUpdate
        '//�R�}���h�E�{�^������
        Call pLockedControl(True)
        Exit Sub
    End If
    
    Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] �̃}�X�^���f���J�n����܂����B")
    
    On Error GoTo cmdUpdate_ClickError:
    Call gdDBS.Database.BeginTrans
    
#If ORA_DEBUG = 1 Then
    Dim updDyn As OraDynaset, recCnt As Long
#Else
    Dim updDyn As Object, recCnt As Long
#End If
    Dim ms As New MouseClass
    Call ms.Start
    
    sql = "SELECT a.*" & vbCrLf
    sql = sql & " FROM " & mRimp.TcHogoshaImport & " a " & vbCrLf
    sql = sql & " WHERE 1 = 1" & vbCrLf
    sql = sql & Condition & vbCrLf
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
'//2007/07/19 �����߂茏����\��
    Dim modoriCnt As Long
'//2007/06/11 ��ʂ� AutoLog �ɂ������̂Ńg���K���~
    Call gdDBS.TriggerControl("tcHogoshaMaster", False)
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
        sql = "SELECT b.* "
        sql = sql & " FROM tcHogoshaMaster b "
        sql = sql & " WHERE CAKYCD = " & gdDBS.ColumnDataSet(dyn.Fields("CIKYCD"), vEnd:=True)
        sql = sql & "   AND CAKSCD = " & gdDBS.ColumnDataSet(dyn.Fields("CIKSCD"), vEnd:=True)
        sql = sql & "   AND CAHGCD = " & gdDBS.ColumnDataSet(dyn.Fields("CIHGCD"), vEnd:=True)
        sql = sql & " ORDER BY CASQNO DESC"     '//�ŏI���R�[�h�݂̂��X�V�Ώ�
#If ORA_DEBUG = 1 Then
        Set updDyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
        Set updDyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
        If updDyn.EOF Then
            If False = pHogoshaInsert(dyn) Then
                GoTo cmdUpdate_ClickError:
            End If
        Else
            If False = pHogoshaUpdate(updDyn, dyn) Then
                GoTo cmdUpdate_ClickError:
            End If
            modoriCnt = modoriCnt + 1
        End If
        Call updDyn.Close
        Set updDyn = Nothing
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    Set dyn = Nothing
    '//�}�X�^���f���ɂ�������������̂ŋ��ʉ�
    If pMoveTempRecords(Condition, gcFurikaeImportToMaster) < 0 Then
        GoTo cmdUpdate_ClickError:
    End If
    Call gdDBS.Database.CommitTrans
'//2007/06/11 �擪�Œ�~���Ă���̂Ńg���K���ĊJ
    Call gdDBS.TriggerControl("tcHogoshaMaster")
    
    pgrProgressBar.Max = pgrProgressBar.Max
    fraProgressBar.Visible = False
    
    '//�X�e�[�^�X�s�̐���E����
    stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "���f����"
    Call MsgBox("�}�X�^���f�Ώ� = [" & cboImpDate.Text & "]" & vbCrLf & vbCrLf & _
                recCnt & " �����}�X�^���f����܂���." & vbCrLf & vbCrLf & _
                "���A�����߂�̌����� " & modoriCnt & " ���ł��B", vbOKOnly + vbInformation, mCaption)
    Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] �� " & recCnt & " ���̔��f���������܂����B���A�����߂�̌����� " & modoriCnt & " ���ł��B")
    '//���X�g���Đݒ�
    Call pMakeComboBox
    '//�{�^����߂�
    cmdUpdate.Caption = cBtnUpdate
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
    Exit Sub
cmdUpdate_ClickError:
    Call gdDBS.Database.Rollback
'//2007/06/11 �擪�Œ�~���Ă���̂Ńg���K���ĊJ
    Call gdDBS.TriggerControl("tcHogoshaMaster")
    If err Then
        Dim errCode As Integer, errMsg As String
        If gdDBS.Database.LastServerErr Then
            errCode = gdDBS.Database.LastServerErr
            errMsg = gdDBS.Database.LastServerErrText
        Else
            errCode = err
            errMsg = Error
        End If
        fraProgressBar.Visible = False
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "���f�G���[(" & errCode & ")"
        Call gdDBS.AutoLogOut(mCaption, "�}�X�^���f�Ώ� = [" & cboImpDate.Text & "] �̓G���[�������������߃}�X�^���f�͒��~����܂����B(Error=" & errMsg & ")")
        Call MsgBox("�}�X�^���f�Ώ� = [" & cboImpDate.Text & "]" & vbCrLf & _
                    "�̓G���[�������������߃}�X�^���f�͒��~����܂����B" & vbCrLf & errMsg, _
                vbOKOnly + vbCritical, mCaption)
    Else
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "���f���f"
        Call gdDBS.AutoLogOut(mCaption, "�}�X�^���f�Ώ� = [" & cboImpDate.Text & "]" & vbCrLf & "�̃}�X�^���f�͒��~����܂����B")
    End If
    '//�{�^����߂�
    cmdUpdate.Caption = cBtnUpdate
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
End Sub

Private Sub pMakeComboBox()
    Dim ms As New MouseClass
    Call ms.Start
    '//�R�}���h�E�{�^������
    Call pLockedControl(False)
'    Dim sql As String, dyn As OraDynaset, MaxDay As Variant
    Dim sql As String, dyn As Object, MaxDay As Variant
    sql = "SELECT DISTINCT TO_CHAR(CIINDT,'yyyy/mm/dd hh24:mi:ss') CIINDT_A"
    sql = sql & " FROM " & mRimp.TcHogoshaImport
    sql = sql & " ORDER BY CIINDT_A"
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
    Call cboImpDate.Clear
    Do Until dyn.EOF()
        Call cboImpDate.AddItem(dyn.Fields("CIINDT_A"))
        'cboImpDate.ItemData(cboImpDate.NewIndex) = dyn.Fields("CIINDT_B")
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    If cboImpDate.ListCount Then
        cboImpDate.ListIndex = cboImpDate.ListCount - 1
    Else
        sprMeisai.MaxRows = 0
    End If
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
End Sub

Private Sub Form_Activate()
'    If sprMeisai.ColWidth(eSprCol.eMaxCols) Then
End Sub

Private Sub Form_Load()
    Me.Show
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    Call mSpread.Init(sprMeisai)
    lblModoriCount.Caption = "�y �����߂茏���F " & Format(0, "#,0") & " �� �z"
    lblModoriCount.Refresh
    '//Spread�̗񒲐�
    Dim ix As Long
    With sprMeisai
        Call sprMeisai_LostFocus    '//ToolTip ��ݒ�
        .MaxCols = eSprCol.eMaxCols
        '//�G���[�������̂ŕ\����(eUseCol)�ȍ~�͔�\���ɂ���
        For ix = eSprCol.eUseCols + 1 To eSprCol.eMaxCols
            .ColWidth(ix) = 0
        Next ix
        '.ColWidth(eSprCol.eImpDate) = 0
        '.ColWidth(eSprCol.eImpSEQ) = 0
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
    'Call cmdEnd.SetFocus
End Sub

Private Sub Form_Resize()
    '//����ȏ㏬��������ƃR���g���[�����B���̂Ő��䂷��
    If Me.Height < 8500 Then
        Me.Height = 8500
    End If
    If Me.Width < 11220 Then
        Me.Width = 11220
    End If
    Call mForm.Resize
    fraProgressBar.Left = 1860
    fraProgressBar.Top = Me.Height - 970
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mAbort = True
    Set mForm = Nothing
    Set mReg = Nothing
    If Not gdFormSub Is Nothing Then
        Unload gdFormSub
    End If
    Set gdFormSub = Nothing
    '//�Ō�ɂ��Ȃ��Ƃ��̃t�H�[���̑�����̎Q�Ƃɂ��ă��[�h�����
    Set frmFurikaeReqImport = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Public Sub gEditToSpreadSheet(vMove As Integer)
'// vMove => -1:�O���ړ� / 0:�ړ����� / 1:����ړ�
'CIITKB eItakuName           '�ϑ��Җ�
'CIKYCD eKeiyakuCode         '�_��҃R�[�h
'       eKeiyakuName         '�_��Җ�
'CIKSCD eKyoshitsuNo         '�����ԍ�
'CIHGCD eHogoshaCode         '�ی�҃R�[�h
'CIKJNM eHogoshaName         '�ی�Җ�(����)
'CIKNNM eHogoshaKana         '�ی�Җ�(�J�i)=>�������`�l��
'CISTNM eSeitoName           '���k����
'CISKGK eFurikaeGaku         '�U�֋��z
'CIKKBN eKinyuuKubun         '���Z�@�֋敪
'CIBANK eBankCode            '��s�R�[�h
'       eBankName_m          '��s��(�}�X�^�[)
'CIBKNM eBankName_i          '��s��(�捞)
'CISITN eShitenCode          '�x�X�R�[�h
'       eShitenName_m        '�x�X��(�}�X�^�[)
'CISINM eShitenName_i        '�x�X��(�捞)
'CIKZSB eYokinShumoku        '�a�����
'CIKZNO eKouzaBango          '�����ԍ�
'CIYBTK eYubinKigou          '�X�֋�:�ʒ��L��
'CIYBTN eYubinBango          '�X�֋�:�ʒ��ԍ�
'CIKZNM eKouzaName           '�������`�l=>�ی�Җ�(�J�i)
'CIINDT eImpDate             '�捞��
'CISEQN eImpSEQ              '�r�d�p

    '//�s�̃f�[�^����v���Ă��Ȃ���Βu�������Ȃ�
    If Not (Format(gdFormSub.lblCIINDT.Caption, "yyyy/MM/dd hh:nn:ss") = Format(mSpread.Value(eSprCol.eImpDate, mEditRow), "yyyy/MM/dd hh:nn:ss") _
        And gdFormSub.lblCISEQN.Caption = mSpread.Value(eSprCol.eImpSEQ, mEditRow) _
      ) Then
        Call MsgBox("�s�f�[�^���ُ�Ȉ�" & vbCrLf & "�X�V�o���܂���ł���.", vbOKOnly + vbCritical, mCaption)
        Exit Sub
    End If
    Dim obj As Object
    mSpread.Value(eSprCol.eErrorStts, mEditRow) = cEditDataMsg
    mSpread.BackColor(eSprCol.eErrorStts, mEditRow) = mRimp.ErrorStatus(mRimp.errEditData)
    For Each obj In gdFormSub.Controls
        If TypeOf obj Is imText _
        Or TypeOf obj Is imNumber _
        Or TypeOf obj Is imDate _
        Or TypeOf obj Is Label Then
            '//�R���g���[���� DataChanged �v���p�e�B���������čX�V��K�v�Ƃ��邩���f
            If "" <> obj.DataField And True = obj.DataChanged Then
                Select Case UCase(Right(obj.Name, 6))
                Case "CIITKB" '//eItakuName           '�ϑ��Җ�
                    mSpread.Value(eSprCol.eItakuName, mEditRow) = gdFormSub.cboABKJNM.Text
                Case "CIKYCD" '//eKeiyakuCode         '�_��҃R�[�h
                              '//eKeiyakuName         '�_��Җ�
                    mSpread.Value(eSprCol.eKeiyakuCode, mEditRow) = obj.Text
                    mSpread.Value(eSprCol.eKeiyakuName, mEditRow) = gdFormSub.lblBAKJNM.Caption
                Case "CIKSCD" '//eKyoshitsuNo         '�����ԍ�
                    mSpread.Value(eSprCol.eKyoshitsuNo, mEditRow) = obj.Text
                Case "CIHGCD" '//eHogoshaCode         '�ی�҃R�[�h
                    mSpread.Value(eSprCol.eHogoshaCode, mEditRow) = obj.Text
                Case "CIKJNM" '//eHogoshaName         '�ی�Җ�(����)
                    mSpread.Value(eSprCol.eHogoshaName, mEditRow) = obj.Text
                Case "CIKNNM" '//eHogoshaKana         '�ی�Җ�(�J�i)=>�������`�l��
                    mSpread.Value(eSprCol.eHogoshaKana, mEditRow) = obj.Text
                Case "CISTNM" '//eSeitoName           '���k����
                    mSpread.Value(eSprCol.eSeitoName, mEditRow) = obj.Text
                Case "CISKGK" '//eFurikaeGaku         '�U�֋��z
                    mSpread.Value(eSprCol.eFurikaeGaku, mEditRow) = obj.Text
                Case "CIKKBN" '//eKinyuuKubun         '���Z�@�֋敪
                    If 0 = gdFormSub.lblCIKKBN.Caption Or 1 = gdFormSub.lblCIKKBN.Caption Then
                        mSpread.Value(eSprCol.eKinyuuKubun, mEditRow) = gdFormSub.optCIKKBN(gdFormSub.lblCIKKBN.Caption).Caption
                    End If
                Case "CIBANK" '//eBankCode            '��s�R�[�h
                              '//eBankName_m          '��s��(�}�X�^�[)
                    mSpread.Value(eSprCol.eBankCode, mEditRow) = obj.Text
                    mSpread.Value(eSprCol.eBankName_m, mEditRow) = gdFormSub.lblBankName.Caption
                Case "CIBKNM" '//eBankName_i          '��s��(�捞)
                    mSpread.Value(eSprCol.eBankName_i, mEditRow) = obj.Text
                Case "CISITN" '//eShitenCode          '�x�X�R�[�h
                              '//eShitenName_m        '�x�X��(�}�X�^�[)
                    mSpread.Value(eSprCol.eShitenCode, mEditRow) = obj.Text
                    mSpread.Value(eSprCol.eShitenName_m, mEditRow) = gdFormSub.lblShitenName.Caption
                Case "CISINM" '//eShitenName_i        '�x�X��(�捞)
                    mSpread.Value(eSprCol.eShitenName_i, mEditRow) = obj.Text
                Case "CIKZSB" '//eYokinShumoku        '�a�����
                    If 1 = gdFormSub.lblCIKZSB.Caption Or 2 = gdFormSub.lblCIKZSB.Caption Then
                        mSpread.Value(eSprCol.eYokinShumoku, mEditRow) = gdFormSub.optCIKZSB(gdFormSub.lblCIKZSB.Caption).Caption
                    End If
                Case "CIKZNO" '//eKouzaBango          '�����ԍ�
                    mSpread.Value(eSprCol.eKouzaBango, mEditRow) = obj.Text
                Case "CIYBTK" '//eYubinKigou          '�X�֋�:�ʒ��L��
                    mSpread.Value(eSprCol.eYubinKigou, mEditRow) = obj.Text
                Case "CIYBTN" '//eYubinBango          '�X�֋�:�ʒ��ԍ�
                    mSpread.Value(eSprCol.eYubinBango, mEditRow) = obj.Text
                Case "CIKZNM" '//eKouzaName           '�������`�l=>�ی�Җ�(�J�i)
                    mSpread.Value(eSprCol.eKouzaName, mEditRow) = obj.Text
                End Select
            End If
        End If
    Next obj
    mEditRow = mEditRow + vMove   '//-1:�O���ړ� / 0:�ړ����� / 1:����ړ�
End Sub

Private Sub sprMeisai_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Not gdFormSub Is Nothing Then
        '//�����Ȃ��H
        'If gdFormSub.dbcImport.EditMode <> OracleConstantModule.ORADATA_EDITNONE Then
            If vbOK <> MsgBox("���ݕҏW���̃f�[�^�͔j������܂�.", vbOKCancel + vbInformation, mCaption) Then
                Exit Sub
            End If
            Call gdFormSub.dbcImport.UpdateControls   '//�L�����Z��
        'End If
        'Unload gdFormSub
    End If
    If Row <= 0 Then
        Exit Sub
    End If
    '//�C����ʂ֓n��
    mEditRow = Row
    Set gdFormSub = frmFurikaeReqImportEdit
    Call gdFormSub.Show
    gdFormSub.dbcImportEdit.RecordSource = "SELECT * " & mMainSQL
    Call gdFormSub.dbcImportEdit.Refresh
    Call gdFormSub.dbcImportEdit.Recordset.FindFirst( _
            "     CIINDT = TO_DATE('" & Format(mSpread.Value(eSprCol.eImpDate, Row), "yyyy/MM/dd hh:nn:ss") & "','yyyy/mm/dd hh24:mi:ss') " & _
            " AND CISEQN = " & mSpread.Value(eSprCol.eImpSEQ, Row))
    Call gdFormSub.txtCIKYCD_KeyDown(vbKeyReturn, 0)    '//�_��Җ��������\��
End Sub

Private Sub sprMeisai_LostFocus()
    With sprMeisai
        .TextTipDelay = 1
        .TextTip = TextTipFixedFocusOnly
        .ToolTipText = "�N���b�N�����" & vbCrLf & "�u�捞���`�F�b�N�̏������ʁv��" & vbCrLf & "�ڍׂ��\������܂�."
    End With
End Sub

Private Sub sprMeisai_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    If 0 < Row Then
        sprMeisai.ToolTipText = mSpread.Value(eSprCol.eErrorStts, Row)
        '//�@�\���Ȃ��I
        'sprMeisai.SetTextTipAppearance "�l�r �S�V�b�N", 15, True, True, vbBlue, vbWhite
    End If
End Sub

Private Sub sprMeisai_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    '// OldTop = 1 �̎��̓C�x���g���N���Ȃ�
#If True = VIRTUAL_MODE Then
    Call pSpreadSetErrorStatus
#Else
    If OldTop <> NewTop Then     '//���ׂăo�b�t�@�ɂ���̂őO�s�ɖ߂鎞�͂��Ȃ��悤��
        Call pSpreadSetErrorStatus
    End If
#End If
End Sub

'//�Z���P�ʂɃG���[�ӏ����J���[�\��
Private Sub pSpreadSetErrorStatus(Optional vReset As Boolean = False)
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim ErrStts() As Variant, ix As Integer, cnt As Long
    Dim ms As New MouseClass
    Call ms.Start
'    eErrorStts = 1  '�G���[���e�F�ُ�A����A�x��
'    eItakuName      '�ϑ��Җ�
'    eKeiyakuCode    '�_��҃R�[�h
'    eKeiyakuName    '�_��Җ�
'    eKyoshitsuNo    '�����ԍ�
'    eHogoshaCode    '�ی�҃R�[�h
'    eHogoshaName    '�ی�Җ�(����)
'    eHogoshaKana    '�ی�Җ�(�J�i)=>�������`�l��
'    eSeitoName      '���k����
'    eFurikaeGaku    '�U�֋��z
'    eKinyuuKubun    '���Z�@�֋敪
'    eBankCode       '��s�R�[�h
'    eBankName_m     '��s��(�}�X�^�[)
'    eBankName_i     '��s��(�捞)
'    eShitenCode     '�x�X�R�[�h
'    eShitenName_m   '�x�X��(�}�X�^�[)
'    eShitenName_i   '�x�X��(�捞)
'    eYokinShumoku     '�������
'    eKouzaBango     '�����ԍ�
'    eYubinKigou     '�X�֋�:�ʒ��L��
'    eYubinBango     '�X�֋�:�ʒ��ԍ�
'    eKouzaName      '�������`�l=>�ی�Җ�(�J�i)
    
    If sprMeisai.MaxRows = 0 Then
        Exit Sub
    End If
    '//�R�}���h�E�{�^������
    Call pLockedControl(False)
    '//�G���[���ݒ�
    ErrStts = Array("CIERROr", "CIITKBe", _
                    "CIKYCDe", "cikycde", "CIKSCDe", "CIHGCDe", "CIKJNMe", "CIKNNMe", "CISTNMe", "CISKGKe", _
                    "CIKKBNe", "CIBANKe", "cibanke", "CIBKNMe", "CISITNe", "cisitne", "CISINMe", "CIKZSBe", "CIKZNOe", _
                    "CIYBTKe", "CIYBTNe", _
                    "CIKZNMe" _
                )
    sql = "SELECT ROWNUM,a.* FROM(" & vbCrLf
    sql = sql & "SELECT CIINDT,CISEQN,CIMUPD," & mRimp.StatusColumns("," & vbCrLf, Len("," & vbCrLf))
    sql = sql & mMainSQL
    sql = sql & ") a"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If False = vReset Then
        'SPread �̃X�N���[���o�[�������̂݊J�n�s�Ɉړ�
        Call dyn.FindFirst("ROWNUM >= " & sprMeisai.TopRow)
    End If
    mSpread.Redraw = False
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
            mSpread.BackColor(ix + 1, dyn.RowPosition) = mRimp.ErrorStatus(dyn.Fields(ErrStts(ix)))
        Next ix
        '//�������ʗ�̕\���F
        '//2006/04/04 �}�X�^���f�n�j�t���O���f
        If 0 <> Val(dyn.Fields("CIMUPD")) Then
            mSpread.BackColor(eSprCol.eErrorStts, dyn.RowPosition) = vbYellow
        ElseIf mRimp.ErrorStatus(0) = mSpread.BackColor(eSprCol.eErrorStts, dyn.RowPosition) Then
            mSpread.BackColor(eSprCol.eErrorStts, dyn.RowPosition) = vbCyan
        End If
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    Set dyn = Nothing
    mSpread.Redraw = True
    '//�R�}���h�E�{�^������
    Call pLockedControl(True)
End Sub



