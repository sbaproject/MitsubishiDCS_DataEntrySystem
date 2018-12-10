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

Private mTable As Integer  '//0=契約者 / 1=口座振替(保護者)

Private mYubinCode As String
Private mYubinName As String

Private mKinyuKikan_Fields As Variant
Private mBank_______Fields As Variant
Private mShiten_____Fields As Variant
Private mShubetsu___Fields As Variant
Private mKouzaNo____Fields As Variant
Private mTsutyoNo___Fields As Variant
Private mTsutyoKigouFields As Variant

Private Type tpKeiyakusha   '//契約者
    a001 As String * 5  '委託者番号                 N   5
    a002 As String * 5  '契約者番号（教室）         N   5
    a003 As String * 3  '教室区分                   N   3
'//2002/11/26 ホスト渡しはここに BLANK(5) を追加
    a004 As String * 4  '銀行コード                 N   4
    a005 As String * 3  '支店コード                 N   3
    a006 As String * 1  '預金種目                   N   1
    a007 As String * 7  '口座番号                   N   7
    a008 As String * 30 '口座名義人名（カナ）       C   30
    a009 As String * 3  '郵便番号１                 N   3
    a010 As String * 4  '郵便番号２                 N   4
    a011 As String * 50 '住所１（漢字）             C   50
    a012 As String * 50 '住所２（漢字）             C   50
    a013 As String * 50 '住所３（漢字）             C   50
    a014 As String * 20 '氏名                       C   20
    a015 As String * 7  '教室番号                   N   7
    a016 As String * 13 '電話番号１                 C   13  '//2002/12/09 携帯番号が１３桁要すので１２桁から変更:ホスト使用違いによる変更(DCS芳地氏依頼による)
    a017 As String * 13 '電話番号2                  C   13
    a018 As String * 13 '電話番号3                  C   13
    a019 As String * 13 'ＦＡＸ番号１               C   13
    a020 As String * 13 'ＦＡＸ番号２               C   13
    a021 As String * 3  '新規人数                   N   3
    a022 As String * 6  '調整額                     N   6
End Type

Private Type tpKouzaFurikae '//口座振替=保護者
    a001 As String * 5  '委託者番号                 N   5
    a002 As String * 5  '契約者番号（教室）         N   5
    a003 As String * 3  '教室区分                   N   3
    a004 As String * 3  'ゼロスペース               N   3
    a005 As String * 4  '保護者番号                 N   4
'//2002/11/26 ホスト渡しはここに BLANK(5) を追加
    a006 As String * 4  '銀行コード                 N   4
    a007 As String * 3  '支店コード                 N   3
    a008 As String * 1  '預金種目                   N   1
    a009 As String * 7  '口座番号                   N   7
    a010 As String * 30 '口座名義人名（カナ）       C   30
    a011 As String * 7  '引落金額                   N   7
    a012 As String * 1  '新規コード                 N   1
End Type

Private mKeiyakusha         As tpKeiyakusha     '契約者
Private mHogosha            As tpKouzaFurikae   '保護者=口座振替
Private mKouzaFurikae       As tpKouzaFurikae   '口座振替

Private mLength    As Variant
Private mLen_A     As Variant
Private mLen_B     As Variant
Private mLen_C     As Variant

Private mAttrib    As Variant
Private mAtr_A     As Variant
Private mAtr_B     As Variant
Private mAtr_C     As Variant

Private Enum eType
    Kanji = -2         '全角文字タイプ
    Char = -1          '文字タイプ
    Numeric = 0        '数値タイプ
    Decmal1 = 1        '数値タイプ:小数点以下１桁有る
    Decmal2 = 2        '数値タイプ:小数点以下２桁有る
    Decmal5 = 5        '数値タイプ:小数点以下５桁有る
End Enum

'//契約者
Public Property Get Keiyakusha() As Integer:    Keiyakusha = 0:     End Property
'//保護者
Public Property Get Hogosha() As Integer:       Hogosha = 1:        End Property
'//口座振替
Public Property Get KouzaFurikae() As Integer:  KouzaFurikae = 2:   End Property
'//整数タイプ
Public Property Get N() As Integer: N = eType.Numeric:  End Property
'//文字タイプ
Public Property Get C() As Integer: C = eType.Char:     End Property
'//漢字タイプ
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
    Case N  '// ０フォーマットして右の有効桁分取得
        Temp = Right(Format(Val(IIf(IsNull(vData), "", vData)), String(Length(vField), "0")), Length(vField))
    Case J
        '//文字項目で５１２文字以上のフィールドはない！？
        Temp = StrConv(LeftB(StrConv(Trim(IIf(IsNull(vData), "", vData)) & Space(512), vbFromUnicode), Length(vField)), vbUnicode)
        'Temp = Left(StrConv(Trim(IIf(IsNull(vData), "", vData)) & Space(512), vbWide), Length(vField))
    End Select
    SetData = Temp
End Function

Private Sub Class_Initialize()
    With mKeiyakusha     '契約者
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

    With mHogosha   '保護者=口座振替
        mLen_B = Array( _
                Len(.a001), Len(.a002), Len(.a003), Len(.a004), Len(.a005), Len(.a006), Len(.a007), Len(.a008), Len(.a009), Len(.a010), _
                Len(.a011), Len(.a012) _
            )
        mAtr_B = Array( _
                N, N, N, N, N, N, N, N, N, C, _
                N, N _
            )
    End With

    With mKouzaFurikae   '口座振替
        mLen_C = Array( _
                Len(.a001), Len(.a002), Len(.a003), Len(.a004), Len(.a005), Len(.a006), Len(.a007), Len(.a008), Len(.a009), Len(.a010), _
                Len(.a011), Len(.a012) _
            )
        mAtr_C = Array( _
                N, N, N, N, N, N, N, N, N, C, _
                N, N _
            )
    End With
    '//金融機関を取得する際の統一のために設定する
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
    '//2003/01/31 視認性をよくするために変数名を変更
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
    '//銀行・郵便局は自動的に選択される
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
    '//銀行・郵便局は自動的に選択される
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
    '//銀行・郵便局は自動的に選択される
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
    '//銀行・郵便局は自動的に選択される
    Select Case vDyn.Fields(mKinyuKikan_Fields(mTable))
    Case eBankKubun.KinnyuuKikan
        KouzaNo = vDyn.Fields(mKouzaNo____Fields(mTable))
    Case eBankKubun.YuubinKyoku
'//郵便局は後ろ１桁カット 運用で８桁を必ず入力
        KouzaNo = Left(vDyn.Fields(mTsutyoNo___Fields(mTable)), gcTsuchoBangoMinLen)
    End Select
End Property



