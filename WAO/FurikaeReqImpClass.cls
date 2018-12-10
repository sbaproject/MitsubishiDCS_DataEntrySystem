VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FurikaeReqImpClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
Private Const pcTcHogoshaImport As String = "tcHogoshaImport"
Private mColumns() As Variant
Public mUpdateMessage As Variant

Private Enum eError     '// ??ERROR �̌���
    eImport = -3
    eEditData = -2
    eInvalid = -1
    eNormal = 0
    eWarning = 1
End Enum

Private Enum eUpdate    '// ??OKFG �̌���
    eMin = -2           '//
    eInvalid = -2       '//���f�s�\�F???e �̃t�B�[���h�ɂP���ڂł��u�|�P�v������ꍇ�͔��f�s��
    eWarnErr = -1       '//�x���Ń}�X�^���f���Ȃ�
    eNormal = 0         '//����f�[�^
    eWarnUpd = 1        '//�x���𖳎����Ĕ��f
    eResetCancel = 2    '//���t���O�����Z�b�g���ď㏑���X�V
    eMax = 2            '//
End Enum

#If ORA_DEBUG = 1 Then
Private mDyn As OraDynaset
#Else
Private mDyn As Object
#End If

Public Property Get YubinKigouLength() As String
    YubinKigouLength = gcTsuchoKigoMinLen
End Property

Public Property Get YubinBangoLength() As String
'//��`�͂V���ƂȂ��Ă���̂Ł{�P
    YubinBangoLength = gcTsuchoBangoMinLen + 1
End Property

Public Property Get StatusColumns(Optional vAddString As Variant = "", Optional vLastStringCut As Integer = 0) As Variant
    Dim ix As Integer
    For ix = LBound(mColumns) To UBound(mColumns)
        StatusColumns = StatusColumns & mColumns(ix) & vAddString
    Next ix
    If 0 < vLastStringCut Then
        StatusColumns = Left(StatusColumns, Len(StatusColumns) - vLastStringCut)
    End If
End Property

Public Property Get TcHogoshaImport() As String: TcHogoshaImport = pcTcHogoshaImport: End Property

Public Property Set Recordset(vDyn As Object)
    Set mDyn = vDyn
End Property

Public Sub UpdateComboBox(vComboBox As ComboBox, Optional vError As Integer = eError.eNormal, Optional vCancel As Boolean = False)
    Call vComboBox.Clear
    '// eUpdate �̓��e���R���{�{�b�N�X�ɐݒ肷��
    'Private Enum eUpdate �̓��e���R���{�{�b�N�X�ɐݒ肷��
    '   eInvalid = -2       '//���f�s�\�F???e �̃t�B�[���h�ɂP���ڂł��u�|�P�v������ꍇ�͔��f�s��
    '   eWarnErr = -1       '//�x���Ń}�X�^���f���Ȃ�
    '   eNormal = 0         '//����f�[�^
    '   eWarnUpd = 1        '//�x���𖳎����Ĕ��f
    '   eResetCancel = 2    '//���t���O�����Z�b�g���ď㏑���X�V
    'End Enum
    Dim ix As Integer
    For ix = eUpdate.eMin To eUpdate.eMax
        Call vComboBox.AddItem(mUpdateMessage(ix))
        vComboBox.ItemData(vComboBox.NewIndex) = ix
    Next ix
End Sub

Public Property Get errImport() As Integer:     errImport = eError.eImport:     End Property
Public Property Get errEditData() As Integer:   errEditData = eError.eEditData: End Property
Public Property Get errInvalid() As Integer:    errInvalid = eError.eInvalid:   End Property
Public Property Get errWarning() As Integer:    errWarning = eError.eWarning:   End Property
Public Property Get errNormal() As Integer:     errNormal = eError.eNormal:     End Property

Public Property Get updInvalid() As Integer:            updInvalid = eUpdate.eInvalid:                      End Property
Public Property Get updWarnErr() As Integer:            updWarnErr = eUpdate.eWarnErr:                      End Property
Public Property Get updNormal() As Integer:             updNormal = eUpdate.eNormal:                        End Property
Public Property Get updWarnUpd() As Integer:            updWarnUpd = eUpdate.eWarnUpd:                      End Property
Public Property Get updResetCancel() As Integer:        updResetCancel = eUpdate.eResetCancel:              End Property

'//�G���[�ɂ���ĕ\���F��ύX
Public Property Get ErrorStatus(vData As Variant, Optional vTextBox As Boolean = True) As Long
    Select Case vData
    Case eError.eEditData, eError.eImport  '//�C���f�[�^�A�捞����f�[�^
        ErrorStatus = vbMagenta
    Case eError.eInvalid  '//�G���[
        ErrorStatus = vbRed
    Case eError.eNormal   '//����
        ErrorStatus = IIf(vTextBox, vbWindowBackground, vbButtonFace)
    Case eError.eWarning  '//���[�j���O
        ErrorStatus = vbYellow
    End Select
End Property

Public Property Get ItakushaKubun(Optional vItakuCode As String = "", Optional vKeiyakuCode As String = "") As Variant
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    sql = "SELECT ABITKB "
    sql = sql & " FROM taItakushaMaster"
    sql = sql & " WHERE 1 = 1"  '//���܂��Ȃ�
    If "" <> vItakuCode Then
        sql = sql & " AND ABITKB = '" & vItakuCode & "'"
    ElseIf "" <> vKeiyakuCode Then
        sql = sql & " AND ABKYTP = '" & Left(vKeiyakuCode, 1) & "'"
    Else
        sql = sql & " AND 1 = -1"   '//���ʃZ�b�g�͕K�� EOF()
    End If
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
    If "" <> vItakuCode Then
        ItakushaKubun = Not dyn.EOF
    ElseIf "" <> vKeiyakuCode Then
        If Not dyn.EOF Then
            ItakushaKubun = dyn.Fields("ABITKB")
        End If
    End If
    Call dyn.Close
    Set dyn = Nothing
End Property

Public Property Get KeiyakushaCode(vKeiyakuCode As String) As Variant
    KeiyakushaCode = Not IsEmpty(ItakushaKubun(vKeiyakuCode:=vKeiyakuCode))
End Property

Private Sub Class_Initialize()
    '/////////////////////////////////////////////////////////////////
    '//�G���[�t���O��擾�F�񖼂U�����ȏ�̓G���[�E�X�e�[�^�X�̍��� //
    '/////////////////////////////////////////////////////////////////
    mColumns = gdDBS.FieldNames(pcTcHogoshaImport, " AND LENGTH(COLUMN_NAME) > 6")
    ReDim mUpdateMessage(eUpdate.eMin To eUpdate.eMax) As Variant
    mUpdateMessage(eUpdate.eInvalid) = "�� ���f�͕s�\(�ُ�f�[�^)"
    mUpdateMessage(eUpdate.eWarnErr) = "�� ���f�͕s�\(�x���f�[�^)"
    mUpdateMessage(eUpdate.eNormal) = "�� ���f�͉\(����f�[�^)"
    mUpdateMessage(eUpdate.eWarnUpd) = "�� ���f�͉\(�x���𖳎�)"
    mUpdateMessage(eUpdate.eResetCancel) = "�� ���f�͉\(��������)"
End Sub
