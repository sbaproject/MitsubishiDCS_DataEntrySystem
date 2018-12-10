VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StructureClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mTable As Integer  '//0=�_��� / 1=�����U��(�ی��)

Private mYubinCode As String
Private mYubinName As String

Private mKinyuKikan_Fields As Variant
Private mBank_______Fields As Variant
Private mShiten_____Fields As Variant
Private mShubetsu___Fields As Variant
Private mKouzaNo____Fields As Variant
Private mTsutyoNo___Fields As Variant
Private mTsutyoKigouFields As Variant

Private Type tpKeiyakusha   '//�_���
    a001 As String * 5  '�ϑ��Ҕԍ�                 N   5
    a002 As String * 5  '�_��Ҕԍ��i�����j         N   5
    a003 As String * 3  '�����敪                   N   3
'//2002/11/26 �z�X�g�n���͂����� BLANK(5) ��ǉ�
    a004 As String * 4  '��s�R�[�h                 N   4
    a005 As String * 3  '�x�X�R�[�h                 N   3
    a006 As String * 1  '�a�����                   N   1
    a007 As String * 7  '�����ԍ�                   N   7
    a008 As String * 30 '�������`�l���i�J�i�j       C   30
    a009 As String * 3  '�X�֔ԍ��P                 N   3
    a010 As String * 4  '�X�֔ԍ��Q                 N   4
    a011 As String * 50 '�Z���P�i�����j             C   50
    a012 As String * 50 '�Z���Q�i�����j             C   50
    a013 As String * 50 '�Z���R�i�����j             C   50
    a014 As String * 20 '����                       C   20
    a015 As String * 7  '�����ԍ�                   N   7
    a016 As String * 13 '�d�b�ԍ��P                 C   13  '//2002/12/09 �g�єԍ����P�R���v���̂łP�Q������ύX:�z�X�g�g�p�Ⴂ�ɂ��ύX(DCS�F�n���˗��ɂ��)
    a017 As String * 13 '�d�b�ԍ�2                  C   13
    a018 As String * 13 '�d�b�ԍ�3                  C   13
    a019 As String * 13 '�e�`�w�ԍ��P               C   13
    a020 As String * 13 '�e�`�w�ԍ��Q               C   13
    a021 As String * 3  '�V�K�l��                   N   3
    a022 As String * 6  '�����z                     N   6
End Type

Private Type tpKouzaFurikae '//�����U��=�ی��
    a001 As String * 5  '�ϑ��Ҕԍ�                 N   5
    a002 As String * 5  '�_��Ҕԍ��i�����j         N   5
    a003 As String * 3  '�����敪                   N   3
    a004 As String * 3  '�[���X�y�[�X               N   3
    a005 As String * 4  '�ی�Ҕԍ�                 N   4
'//2002/11/26 �z�X�g�n���͂����� BLANK(5) ��ǉ�
    a006 As String * 4  '��s�R�[�h                 N   4
    a007 As String * 3  '�x�X�R�[�h                 N   3
    a008 As String * 1  '�a�����                   N   1
    a009 As String * 7  '�����ԍ�                   N   7
    a010 As String * 30 '�������`�l���i�J�i�j       C   30
    a011 As String * 7  '�������z                   N   7
    a012 As String * 1  '�V�K�R�[�h                 N   1
End Type

Private mKeiyakusha         As tpKeiyakusha     '�_���
Private mHogosha            As tpKouzaFurikae   '�ی��=�����U��
Private mKouzaFurikae       As tpKouzaFurikae   '�����U��

Private mLength    As Variant
Private mLen_A     As Variant
Private mLen_B     As Variant
Private mLen_C     As Variant

Private mAttrib    As Variant
Private mAtr_A     As Variant
Private mAtr_B     As Variant
Private mAtr_C     As Variant

Private Enum eType
    Kanji = -2         '�S�p�����^�C�v
    Char = -1          '�����^�C�v
    Numeric = 0        '���l�^�C�v
    Decmal1 = 1        '���l�^�C�v:�����_�ȉ��P���L��
    Decmal2 = 2        '���l�^�C�v:�����_�ȉ��Q���L��
    Decmal5 = 5        '���l�^�C�v:�����_�ȉ��T���L��
End Enum

'//�_���
Public Property Get Keiyakusha() As Integer:    Keiyakusha = 0:     End Property
'//�ی��
Public Property Get Hogosha() As Integer:       Hogosha = 1:        End Property
'//�����U��
Public Property Get KouzaFurikae() As Integer:  KouzaFurikae = 2:   End Property
'//�����^�C�v
Public Property Get N() As Integer: N = eType.Numeric:  End Property
'//�����^�C�v
Public Property Get C() As Integer: C = eType.Char:     End Property
'//�����^�C�v
Public Property Get J() As Integer: J = eType.Kanji:    End Property

Public Sub SelectStructure(ByVal vSelect As Integer)
    Select Case vSelect
    Case Keiyakusha
        mLength = mLen_A
        mAttrib = mAtr_A
    Case Hogosha
        mLength = mLen_B
        mAttrib = mAtr_B
    Case KouzaFurikae
        mLength = mLen_C
        mAttrib = mAtr_C
    End Select
    mTable = vSelect
End Sub

Public Property Get Attrib(ByVal vField As Integer) As Integer
    Attrib = mAttrib(vField)
End Property

Public Property Get Length(ByVal vField As Integer) As Integer
    Length = mLength(vField)
End Property

Public Function SetData(ByVal vData As Variant, ByVal vField As Integer) As String
    Dim Temp As String
    Select Case mAttrib(vField)
    Case C
        Temp = Left(Trim(IIf(IsNull(vData), "", vData)) & Space(Length(vField)), Length(vField))
    Case N  '// �O�t�H�[�}�b�g���ĉE�̗L�������擾
        Temp = Right(Format(Val(IIf(IsNull(vData), "", vData)), String(Length(vField), "0")), Length(vField))
    Case J
        '//�������ڂłT�P�Q�����ȏ�̃t�B�[���h�͂Ȃ��I�H
        Temp = StrConv(LeftB(StrConv(Trim(IIf(IsNull(vData), "", vData)) & Space(512), vbFromUnicode), Length(vField)), vbUnicode)
        'Temp = Left(StrConv(Trim(IIf(IsNull(vData), "", vData)) & Space(512), vbWide), Length(vField))
    End Select
    SetData = Temp
End Function

Private Sub Class_Initialize()
    With mKeiyakusha     '�_���
        mLen_A = Array( _
                Len(.a001), Len(.a002), Len(.a003), Len(.a004), Len(.a005), Len(.a006), Len(.a007), Len(.a008), Len(.a009), Len(.a010), _
                Len(.a011), Len(.a012), Len(.a013), Len(.a014), Len(.a015), Len(.a016), Len(.a017), Len(.a018), Len(.a019), Len(.a020), _
                Len(.a021), Len(.a022) _
            )
        mAtr_A = Array( _
                N, N, N, N, N, N, N, C, N, N, _
                J, J, J, J, N, C, C, C, C, C, _
                N, N _
            )
    End With

    With mHogosha   '�ی��=�����U��
        mLen_B = Array( _
                Len(.a001), Len(.a002), Len(.a003), Len(.a004), Len(.a005), Len(.a006), Len(.a007), Len(.a008), Len(.a009), Len(.a010), _
                Len(.a011), Len(.a012) _
            )
        mAtr_B = Array( _
                N, N, N, N, N, N, N, N, N, C, _
                N, N _
            )
    End With

    With mKouzaFurikae   '�����U��
        mLen_C = Array( _
                Len(.a001), Len(.a002), Len(.a003), Len(.a004), Len(.a005), Len(.a006), Len(.a007), Len(.a008), Len(.a009), Len(.a010), _
                Len(.a011), Len(.a012) _
            )
        mAtr_C = Array( _
                N, N, N, N, N, N, N, N, N, C, _
                N, N _
            )
    End With
    '//���Z�@�ւ��擾����ۂ̓���̂��߂ɐݒ肷��
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
'    Set dyn = gdDBS.OpenRecordset("SELECT * FROM taSystemInformation WHERE AASKEY = '" & gdDBS.SystemKey & "'", dynOption.ORADYN_READONLY)
    Set dyn = gdDBS.OpenRecordset("SELECT * FROM taSystemInformation WHERE AASKEY = '" & gdDBS.SystemKey & "'", OracleConstantModule.ORADYN_READONLY)
    If Not dyn.EOF() Then
        mYubinCode = dyn.Fields("AAYSNO")
        mYubinName = dyn.Fields("AAYSNM")
    End If
    Call dyn.Close
    mTable = -1
    '//2003/01/31 ���F�����悭���邽�߂ɕϐ�����ύX
    mKinyuKikan_Fields = Array("BAKKBN", "CAKKBN", "FAKKBN")
    mBank_______Fields = Array("BABANK", "CABANK", "FABANK")
    mShiten_____Fields = Array("BASITN", "CASITN", "FASITN")
    mShubetsu___Fields = Array("BAKZSB", "CAKZSB", "FAKZSB")
    mKouzaNo____Fields = Array("BAKZNO", "CAKZNO", "FAKZNO")
    mTsutyoKigouFields = Array("BAYBTK", "CAYBTK", "FAYBTK")
    mTsutyoNo___Fields = Array("BAYBTN", "CAYBTN", "FAYBTN")
End Sub

#If ORA_DEBUG = 1 Then
    Public Property Get BankCode(vDyn As OraDynaset) As String
#Else
    Public Property Get BankCode(vDyn As Object) As String
#End If
    '//��s�E�X�֋ǂ͎����I�ɑI�������
    Select Case vDyn.Fields(mKinyuKikan_Fields(mTable))
    Case eBankKubun.KinnyuuKikan
        BankCode = vDyn.Fields(mBank_______Fields(mTable))
    Case eBankKubun.YuubinKyoku
        BankCode = mYubinCode
    End Select
End Property

#If ORA_DEBUG = 1 Then
Public Property Get ShitenCode(vDyn As OraDynaset) As String
#Else
Public Property Get ShitenCode(vDyn As Object) As String
#End If
    '//��s�E�X�֋ǂ͎����I�ɑI�������
    Select Case vDyn.Fields(mKinyuKikan_Fields(mTable))
    Case eBankKubun.KinnyuuKikan
        ShitenCode = vDyn.Fields(mShiten_____Fields(mTable))
    Case eBankKubun.YuubinKyoku
        ShitenCode = vDyn.Fields(mTsutyoKigouFields(mTable))
    End Select
End Property

#If ORA_DEBUG = 1 Then
Public Property Get Shubetsu(vDyn As OraDynaset) As String
#Else
Public Property Get Shubetsu(vDyn As Object) As String
#End If
    '//��s�E�X�֋ǂ͎����I�ɑI�������
    Select Case vDyn.Fields(mKinyuKikan_Fields(mTable))
    Case eBankKubun.KinnyuuKikan
        Shubetsu = vDyn.Fields(mShubetsu___Fields(mTable))
    Case eBankKubun.YuubinKyoku
        Shubetsu = "0"
    End Select
End Property

#If ORA_DEBUG = 1 Then
Public Property Get KouzaNo(vDyn As OraDynaset) As String
#Else
Public Property Get KouzaNo(vDyn As Object) As String
#End If
    Dim tmp As String
    '//��s�E�X�֋ǂ͎����I�ɑI�������
    Select Case vDyn.Fields(mKinyuKikan_Fields(mTable))
    Case eBankKubun.KinnyuuKikan
        KouzaNo = vDyn.Fields(mKouzaNo____Fields(mTable))
    Case eBankKubun.YuubinKyoku
'//�X�֋ǂ͌��P���J�b�g �^�p�łW����K������
        KouzaNo = Left(vDyn.Fields(mTsutyoNo___Fields(mTable)), gcTsuchoBangoMinLen)
    End Select
End Property


