VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptYoteiReqImport 
   Caption         =   "���������s�V�X�e�� - rptYoteiReqImport (ActiveReport)"
   ClientHeight    =   9105
   ClientLeft      =   720
   ClientTop       =   2130
   ClientWidth     =   17490
   WindowState     =   2  '�ő剻
   _ExtentX        =   30850
   _ExtentY        =   16060
   SectionData     =   "�U�֗\��\�捞.dsx":0000
End
Attribute VB_Name = "rptYoteiReqImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mTotalCnt As Long    '//�Ăяo���t�H�[���ŃZ�b�g�����

Private mReport As New ActiveReportClass
Private mYimp As New FurikaeSchImpClass
Private mGroup As String
Private Enum eCount
    eInvalid
    eWarning
    eTotal
End Enum

Private mBodyEdit   As Integer      '//�{�f�B�̕ύX����
Private mBodyGaku   As Currency     '//�{�f�B�̕ύX���z
Private mBodyCancel As Integer      '//�{�f�B�̉�񌏐�

Private mDataCnt(0 To 2) As Long
Private mLineCount As Long

Private Sub ActiveReport_Initialize()
    '//A4 �c�u��
'//2007/06/08 �f�t�H���g�p�����`�S=>�a�S�ɕύX�F�`�S���킴�킴�a�S�ɕύX���ďo�͂��Ă����
    Call mReport.Setup(Me, vbPRPSB4, ddOPortrait)
    'txtShoriDate.Text = Mid(gdADO.SystemData("MASRDT"), 1, 4) & " �N " & Mid(gdADO.SystemData("MASRDT"), 5, 2) & " ���x"
    Me.PageSettings.TopMargin = 500
    Me.PageSettings.LeftMargin = 500
    Me.PageSettings.RightMargin = 500
    Me.PageSettings.BottomMargin = 500
    '//���̎��_�� Load()
    'mDataCnt(eCount.eTotal) = mTotalCnt
    mBodyEdit = 0      '//�{�f�B�̕ύX����
    mBodyGaku = 0      '//�{�f�B�̕ύX���z
    mBodyCancel = 0    '//�{�f�B�̉�񌏐�
End Sub

Private Sub ActiveReport_ReportStart()
    '//�����ł��Ȃ��Ǝ��Ȃ�
'    mDataCnt(eCount.eTotal) = mTotalCnt
    Erase mDataCnt
End Sub

Private Sub Detail_BeforePrint()
'//���̈ʒu�Ń}�X�N���Ȃ��Ƃ��܂��o���Ȃ�
    mLineCount = mLineCount + 1
    'shpMask.BackStyle = IIf(mLineCount Mod 2, ddBKTransparent, ddBKNormal)
End Sub

Private Sub pTextBoxColor(vObj As Object, vStatus As Variant)
    Select Case vStatus
    Case mYimp.errInvalid
        vObj.ForeColor = vbRed
    Case mYimp.errWarning
        vObj.ForeColor = vbMagenta
    Case mYimp.errEditData
        vObj.ForeColor = vbGreen
    Case Else
        vObj.ForeColor = vbBlack
    End Select
    vObj.Font.Underline = vStatus <> 0
End Sub

Private Sub Detail_Format()
    If Me.Fields("FIRKBN") = mYimp.RecordIsTotal Then
        Me.txtFIHKCT_BODY.Text = mBodyEdit & " ��"      '//�{�f�B�̕ύX���� �v�Z����
        Me.txtFIHKKG.Text = Format(mBodyGaku, "#,##0")  '//�{�f�B�̕ύX���z �v�Z����
        Me.txtFIKYFG.Text = mBodyCancel & " ��"         '//�{�f�B�̉�񌏐� �v�Z����
        Me.txtFIHKCT_TOTAL.Text = Me.txtFIHKCT_TOTAL.Text & " ��"
        Me.txtFIKYCT_TOTAL.Text = Me.txtFIKYCT_TOTAL.Text & " ��"
        mBodyEdit = 0      '//�{�f�B�̕ύX����
        mBodyGaku = 0      '//�{�f�B�̕ύX���z
        mBodyCancel = 0    '//�{�f�B�̉�񌏐�
        shpMask.Visible = True
    Else
        '//���v���̕\����}�~
        Me.txtFIHKCT_BODY.Text = ""
        Me.txtFIHKCT_TOTAL.Text = ""
        Me.txtFIHKKG_TOTAL.Text = ""
        Me.txtFIKYCT_TOTAL.Text = ""
        If Not IsNull(Me.Fields("FIHKKG")) Then '//�{�f�B�̕ύX���z
            mBodyGaku = mBodyGaku + gdDBS.Nz(Me.Fields("FIHKKG")) '//�{�f�B�̕ύX���z
        End If
        If 0 <> gdDBS.Nz(Me.Fields("FIKYFG"), 0) Then
            mBodyCancel = mBodyCancel + 1           '//�{�f�B�̉�񌏐�
        Else
            mBodyEdit = mBodyEdit + 1                   '//�{�f�B�̕ύX����
        End If
        shpMask.Visible = False
    End If
    '//sql = sql & "FIITKB || FIKYCD || FIKSCD || FIPGNO || FIFKDT FIGROUP" & vbCrLf
    If mGroup = Me.Fields("FIGROUP").Value Then
        '//2006/04/05 ���ꂼ��ɃG���[��������΃O���[�s���O�ׂ̈ɏ���
        If Me.Fields("FiITKBe").Value = mYimp.errNormal Then
            Me.txtABITCD.Text = ""
        End If
        If Me.Fields("FiKYCDe").Value = mYimp.errNormal Then
            Me.txtFIKYCD.Text = ""
        End If
        If Me.Fields("FiKSCDe").Value = mYimp.errNormal Then
            Me.txtFIKSCD.Text = ""
        End If
        If Me.Fields("FiPGNOe").Value = mYimp.errNormal Then
            Me.txtFIPGNO.Text = ""
        End If
        If Me.Fields("FiFKDTe").Value = mYimp.errNormal Then
            Me.txtFIFKDT.Text = ""
        End If
    Else
        mGroup = Me.Fields("FIGROUP").Value
    End If
    '//�e�������v�Z
    mDataCnt(eCount.eTotal) = mDataCnt(eCount.eTotal) + 1
    Select Case Me.Fields("FiERROr")
    Case mYimp.errInvalid, mYimp.errEditData, mYimp.errImport
        mDataCnt(eCount.eInvalid) = mDataCnt(eCount.eInvalid) + 1
    Case mYimp.errWarning
        mDataCnt(eCount.eWarning) = mDataCnt(eCount.eWarning) + 1
    End Select
    Call pTextBoxColor(txtFIERRNM, Me.Fields("FiERROr").Value)
    Call pTextBoxColor(txtFIFKDT, Me.Fields("FiFKDTe").Value)
    Call pTextBoxColor(txtABITCD, Me.Fields("FiITKBe").Value)
    Call pTextBoxColor(txtFIKYCD, Me.Fields("FiKYCDe").Value)
    Call pTextBoxColor(txtFIKSCD, Me.Fields("FiKSCDe").Value)
    Call pTextBoxColor(txtFIPGNO, Me.Fields("FiPGNOe").Value)
    Call pTextBoxColor(txtFIHGCD, Me.Fields("FiHGCDe").Value)
    Call pTextBoxColor(txtFIHKKG, Me.Fields("FiHKKGe").Value)
    Call pTextBoxColor(txtFIKYFG, Me.Fields("FiKYFGe").Value)
    If Me.Fields("FiHKCTe").Value Then
        Call pTextBoxColor(txtFIHKCT_BODY, Me.Fields("FiHKCTe").Value)
    End If
    Call pTextBoxColor(txtFIHKCT_TOTAL, Me.Fields("FiHKCTe").Value)
    Call pTextBoxColor(txtFIHKKG_TOTAL, Me.Fields("FiHKKGe").Value)
    If Me.Fields("FiHKCTe").Value Then
        Call pTextBoxColor(txtFIKYFG, Me.Fields("FiKYCTe").Value)
    End If
    Call pTextBoxColor(txtFIKYCT_TOTAL, Me.Fields("FiKYCTe").Value)
End Sub

Private Sub PageHeader_Format()
    lblSysDate.Caption = "( " & Format(Now(), "yyyy/mm/dd hh:nn:ss") & " )"
    lblPage.Caption = "Page: " & Me.PageNumber
End Sub

Private Sub ReportFooter_Format()
    '//����f�[�^�͏o�͂��Ȃ��̂ŉ����Z���ĎZ�o
    lblTotalMsg.Caption = ""
    lblTotalMsg.Caption = lblTotalMsg.Caption & " �ُ�F " & Format(mDataCnt(eCount.eInvalid), "#,#0") & " �� / "
    lblTotalMsg.Caption = lblTotalMsg.Caption & " �x���F " & Format(mDataCnt(eCount.eWarning), "#,#0") & " �� / "
    lblTotalMsg.Caption = lblTotalMsg.Caption & " ����F " & Format(mDataCnt(eCount.eTotal) - (mDataCnt(eCount.eInvalid) + mDataCnt(eCount.eWarning)), "#,#0") & " �� / "
    lblTotalMsg.Caption = lblTotalMsg.Caption & " �������F " & Format(mDataCnt(eCount.eTotal), "#,#0") & " ��"
End Sub

