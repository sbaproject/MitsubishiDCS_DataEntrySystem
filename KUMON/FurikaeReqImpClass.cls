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

Private Enum eError     '// ??ERROR の結果
    eImport = -3
    eEditData = -2
    eInvalid = -1
    eNormal = 0
    eWarning = 1
End Enum

Private Enum eUpdate    '// ??OKFG の結果
    eMin = -2           '//
    eInvalid = -2       '//反映不可能：???e のフィールドに１項目でも「－１」がある場合は反映不可
    eWarnErr = -1       '//警告でマスタ反映しない
    eNormal = 0         '//正常データ
    eWarnUpd = 1        '//警告を無視して反映
    eResetCancel = 2    '//解約フラグをリセットして上書き更新
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
'//定義は７桁となっているので＋１
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
    '// eUpdate の内容をコンボボックスに設定する
    'Private Enum eUpdate の内容をコンボボックスに設定する
    '   eInvalid = -2       '//反映不可能：???e のフィールドに１項目でも「－１」がある場合は反映不可
    '   eWarnErr = -1       '//警告でマスタ反映しない
    '   eNormal = 0         '//正常データ
    '   eWarnUpd = 1        '//警告を無視して反映
    '   eResetCancel = 2    '//解約フラグをリセットして上書き更新
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

'//エラーによって表示色を変更
Public Property Get ErrorStatus(vData As Variant, Optional vTextBox As Boolean = True) As Long
    Select Case vData
    Case eError.eEditData, eError.eImport  '//修正データ、取込直後データ
        ErrorStatus = vbMagenta
    Case eError.eInvalid  '//エラー
        ErrorStatus = vbRed
    Case eError.eNormal   '//正常
        ErrorStatus = IIf(vTextBox, vbWindowBackground, vbButtonFace)
    Case eError.eWarning  '//ワーニング
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
    sql = sql & " WHERE 1 = 1"  '//おまじない
    If "" <> vItakuCode Then
        sql = sql & " AND ABITKB = '" & vItakuCode & "'"
    ElseIf "" <> vKeiyakuCode Then
        sql = sql & " AND ABKYTP = '" & Left(vKeiyakuCode, 1) & "'"
    Else
        sql = sql & " AND 1 = -1"   '//結果セットは必ず EOF()
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
    '//エラーフラグ列取得：列名６文字以上はエラー・ステータスの項目 //
    '/////////////////////////////////////////////////////////////////
    mColumns = gdDBS.FieldNames(pcTcHogoshaImport, " AND LENGTH(COLUMN_NAME) > 6")
    ReDim mUpdateMessage(eUpdate.eMin To eUpdate.eMax) As Variant
    mUpdateMessage(eUpdate.eInvalid) = "▲ 反映は不可能(異常データ)"
    mUpdateMessage(eUpdate.eWarnErr) = "● 反映は不可能(警告データ)"
    mUpdateMessage(eUpdate.eNormal) = "◎ 反映は可能(正常データ)"
    mUpdateMessage(eUpdate.eWarnUpd) = "△ 反映は可能(警告を無視)"
    mUpdateMessage(eUpdate.eResetCancel) = "◇ 反映は可能(解約を解除)"
End Sub

