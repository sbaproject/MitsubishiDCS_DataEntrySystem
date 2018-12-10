VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FurikaeSchImpClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
Private Const pcTfFurikaeImport As String = "tfFurikaeYoteiImport"
Private Const pcTotalTextKubun As String = "9"
Private mColumns() As Variant
Public mUpdateMessage As Variant

Private Enum eRecord
    eTotal = -1
'//    eDetail
End Enum

Private Enum eError     '// ??ERROR �̌���
    eDeleted = -4       '//2006/06/16 ���ׂ��폜
    eImport = -3
    eEditData = -2
    eInvalid = -1
    eNormal = 0
    eWarning = 1
End Enum

Private Enum eUpdate    '// ??OKFG �̌���
    eReflection = -5       '//2007/03/12 �}�X�^���f�ς݃f�[�^
    eDeleted = -4       '//2007/03/12 ���ׂ��폜
    eMin = -2           '//
    eInvalid = -2       '//���f�s�\�F???e �̃t�B�[���h�ɂP���ڂł��u�|�P�v������ꍇ�͔��f�s��
    eWarnErr = -1       '//�x���Ń}�X�^���f���Ȃ�
    eNormal = 0         '//����f�[�^
    eWarnUpd = 1        '//�x���𖳎����Ĕ��f
    '//����ȃf�[�^�͖���
    'eResetCancel = 2    '//���t���O�����Z�b�g���ď㏑���X�V
    eMax = 2            '//
End Enum

#If ORA_DEBUG = 1 Then
Private mDyn As OraDynaset
#Else
Private mDyn As Object
#End If

Public Property Get TfFurikaeImport() As String: TfFurikaeImport = pcTfFurikaeImport: End Property

Public Property Get TotalTextKubun() As String: TotalTextKubun = pcTotalTextKubun: End Property

Public Property Get RecordIsTotal() As String:  RecordIsTotal = eRecord.eTotal:  End Property
'Public Property Get RecordIsDetail() As String: RecordIsDetail = eRecord.eDetail: End Property

Public Property Get errDeleted() As Integer:    errDeleted = eError.eDeleted:   End Property
Public Property Get errImport() As Integer:     errImport = eError.eImport:     End Property
Public Property Get errEditData() As Integer:   errEditData = eError.eEditData: End Property
Public Property Get errInvalid() As Integer:    errInvalid = eError.eInvalid:   End Property
Public Property Get errWarning() As Integer:    errWarning = eError.eWarning:   End Property
Public Property Get errNormal() As Integer:     errNormal = eError.eNormal:     End Property

Public Property Get updReflection() As Integer:         updReflection = eUpdate.eReflection:                End Property
Public Property Get updDeleted() As Integer:            updDeleted = eUpdate.eDeleted:                      End Property
Public Property Get updInvalid() As Integer:            updInvalid = eUpdate.eInvalid:                      End Property
Public Property Get updWarnErr() As Integer:            updWarnErr = eUpdate.eWarnErr:                      End Property
Public Property Get updNormal() As Integer:             updNormal = eUpdate.eNormal:                        End Property
Public Property Get updWarnUpd() As Integer:            updWarnUpd = eUpdate.eWarnUpd:                      End Property
'//����ȃf�[�^�͖���
'Public Property Get updResetCancel() As Integer:        updResetCancel = eUpdate.eResetCancel:              End Property

Private Sub Class_Initialize()
    '/////////////////////////////////////////////////////////////////
    '//�G���[�t���O��擾�F�񖼂U�����ȏ�̓G���[�E�X�e�[�^�X�̍��� //
    '/////////////////////////////////////////////////////////////////
    mColumns = gdDBS.FieldNames(TfFurikaeImport, " AND LENGTH(COLUMN_NAME) > 6")
    ReDim mUpdateMessage(eUpdate.eMin To eUpdate.eMax) As Variant
    mUpdateMessage(eUpdate.eInvalid) = "�� ���f�͕s�\(�ُ�f�[�^)"
    mUpdateMessage(eUpdate.eWarnErr) = "�� ���f�͕s�\(�x���f�[�^)"
    mUpdateMessage(eUpdate.eNormal) = "�� ���f�͉\(����f�[�^)"
    mUpdateMessage(eUpdate.eWarnUpd) = "�� ���f�͉\(�x���𖳎�)"
    '//����ȃf�[�^�͖���
    'mUpdateMessage(eUpdate.eResetCancel) = "�� ���f�͉\(��������)"
End Sub

Public Property Get StatusColumns(Optional vAddString As Variant = "", Optional vLastStringCut As Integer = 0) As Variant
    Dim ix As Integer
    For ix = LBound(mColumns) To UBound(mColumns)
        StatusColumns = StatusColumns & mColumns(ix) & vAddString
    Next ix
    If 0 < vLastStringCut Then
        StatusColumns = Left(StatusColumns, Len(StatusColumns) - vLastStringCut)
    End If
End Property

'//�G���[�ɂ���ĕ\���F��ύX
Public Property Get ErrorStatus(vData As Variant, Optional vTextBox As Boolean = True) As Long
    Select Case vData
    Case eError.eEditData, eError.eImport '//�C���f�[�^
        ErrorStatus = vbMagenta
    Case eError.eInvalid  '//�G���[
        ErrorStatus = vbRed
    Case eError.eNormal   '//����
        ErrorStatus = IIf(vTextBox, vbWindowBackground, vbButtonFace)
    Case eError.eWarning  '//���[�j���O
        ErrorStatus = vbYellow
'//2006/06/16 ���׍폜�Ή�
    Case eError.eDeleted
        ErrorStatus = vbGreen
    End Select
End Property
