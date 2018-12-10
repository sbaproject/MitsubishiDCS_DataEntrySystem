VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatabaseClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mReg  As New RegistryClass
Private mConnected As Boolean

#Const ORA_DEBUG = 0
#If ORA_DEBUG = 1 Then
    '//開発時うっとうしいので...。
    Private mOraSSN As OraSessionClass
    Private mOraDBS As OraDatabase
    Private mOraDyn As OraDynaset
#Else
    Private mOraSSN As Object
    Private mOraDBS As Object
    Private mOraDyn As Object
#End If
Private Const cSystemMessage As String = "管理者に報告してください."

Public Function MsgBox(ByVal vMsg As String, ByVal vButton As Integer, ByVal vCap As String) As Integer
    Call AutoLogOut(vCap, vMsg)
    MsgBox = VBA.MsgBox(vMsg, vButton, vCap)
End Function

Public Function AutoLogOut(ByVal vAppl As String, ByVal vMsg As String) As Integer
    If False = mConnected Or False = mReg.zaAutologOut Then
        Exit Function
    End If
    Dim ret As Integer
#If ORA_DEBUG = 1 Then
    Call mOraDBS.Parameters.Add("ret", 0, ORAPARM_INPUT)
    mOraDBS.Parameters("ret").serverType = ORATYPE_NUMBER
#Else
    Call mOraDBS.Parameters.Add("ret", 0, OracleConstantModule.ORAPARM_INPUT)
    mOraDBS.Parameters("ret").serverType = OracleConstantModule.ORATYPE_NUMBER
#End If
    If InStr(vMsg, "'") Then
        vMsg = Replace(vMsg, "'", "''")
    End If
    Call mOraDBS.ExecuteSQL("BEGIN :ret := PKG_UTY.LogOut('" & vAppl & "','" & vMsg & "',0,SYSDATE); END;")
    AutoLogOut = mOraDBS.Parameters("ret").Value
    Call mOraDBS.Parameters.Remove("ret")
End Function

Public Function Holiday(ByVal vYear As Integer) As String
    Dim sql As String
    sql = "SELECT TO_CHAR(TO_DATE(EADATE,'YYYYMMDD'),'MM/DD') AS MMDD"
    sql = sql & " FROM teHolidayMaster"
    sql = sql & " WHERE EADATE LIKE '" & IIf(vYear = -1, Year(Now()), vYear) & "%'"
#If ORA_DEBUG = 1 Then
    Set mOraDyn = Me.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set mOraDyn = Me.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    Do Until mOraDyn.EOF()
'//2003/01/31 Dynaset を Object で定義すると .Value 句を付加しないと Error=5 になる.
        Holiday = Holiday & mOraDyn.Fields("MMDD").Value & ","
        Call mOraDyn.MoveNext
    Loop
    Call mOraDyn.Close
'    If Len(Holiday) Then
'        Holiday = Left(Holiday, Len(Holiday) - 1)
'    End If
End Function

#If ORA_DEBUG = 1 Then
    Public Function SelectBankMaster(vFields As String, vDARKBN As String, vDABANK As String, Optional ByVal vDASITN As String = "", Optional ByVal vDate As Long = 0&) As OraDynaset
#Else
    Public Function SelectBankMaster(vFields As String, vDARKBN As String, vDABANK As String, Optional ByVal vDASITN As String = "", Optional ByVal vDate As Long = 0&) As Object
#End If
    Dim sql As String
    sql = "SELECT " & vFields & " FROM tdBankMaster"
    sql = sql & " WHERE DARKBN = '" & Trim(vDARKBN) & "'"
    sql = sql & "   AND DABANK = '" & Trim(vDABANK) & "'"
    If Trim(vDASITN) <> "" Then
        sql = sql & "   AND DASITN = '" & Trim(vDASITN) & "'"
    End If
'''2002/10/09 ホストデータの関係でフィールドを削除した
'''    '//有効日付を参照する.
'''    If vDate Then
'''        sql = sql & "   AND " & vDate & " BETWEEN DAYKST AND DAYKED"
'''    End If
    'sql = sql & " ORDER BY DARKBN,DABANK,DASITN"
#If ORA_DEBUG = 1 Then
    Set SelectBankMaster = OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set SelectBankMaster = OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
End Function

'----------------------------------------------
'
'   ログインユーザーを取得
'
'----------------------------------------------
Public Property Get LoginUserName() As String
    Dim lRet As Long
    Dim lpReturnedString As String * 256
    Dim nSize As Long
    
    nSize = Len(lpReturnedString)
    lRet = GetUserName(lpReturnedString, nSize)
    LoginUserName = IIf(nSize > 0, Left(lpReturnedString, nSize - 1), "")
End Property

Public Function MoveRecords(vOraDBC As ORADC, Optional ByVal vMove As Integer = 1) As Boolean
    On Error GoTo MoveRecordError
#If ORA_DEBUG = 1 Then
    Dim dyn As OraDynaset
#Else
    Dim dyn As Object
#End If
    Set dyn = vOraDBC.Recordset
    '//先頭レコードよりも前・最後のレコードよりも後への移動は考慮していないので後で考える。
    '// ＋１,－１で移動する分には可能である。
    If vMove > 0 Then
        MoveRecords = Not (True = dyn.IsLast)   '//最後のレコードか？
        If MoveRecords Then
            Call dyn.MoveNextn(Abs(vMove))
        End If
    ElseIf vMove < 0 Then
        MoveRecords = Not (True = dyn.IsFirst)    '//先頭のレコードか？
        If MoveRecords Then
            Call dyn.MovePreviousn(Abs(vMove))
        End If
    End If
MoveRecordError:
    Call dyn.Close
End Function

Public Sub SetItakushaComboBox(ByVal vComboBox As ComboBox)
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset, ix As Integer, def As Integer
#Else
    Dim sql As String, dyn As Object, ix As Integer, def As Integer
#End If

    sql = "SELECT * FROM taItakushaMaster"
    sql = sql & " ORDER BY ABITCD"
#If ORA_DEBUG = 1 Then
    Set dyn = Me.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = Me.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    Call vComboBox.Clear
    Do Until dyn.EOF()
        Call vComboBox.AddItem(dyn.Fields("ABKJNM"), ix)
        vComboBox.ItemData(vComboBox.NewIndex) = dyn.Fields("ABITKB")
        def = IIf(Val(Nz(dyn.Fields("ABDEFF"))) <> 0, ix, def)
        ix = ix + 1
        Call dyn.MoveNext
    Loop
    vComboBox.ListIndex = def
    Call dyn.Close
End Sub

Public Property Get ABITKB() As String
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    sql = "SELECT ABITKB FROM taItakushaMaster"
    sql = sql & " WHERE ABDEFF = 1"
    sql = sql & " ORDER BY ABITCD"
#If ORA_DEBUG = 1 Then
    Set dyn = mOraDBS.CreateDynaset(sql, Options:=dynOption.ORADYN_READONLY)
#Else
    Set dyn = mOraDBS.CreateDynaset(sql, Options:=OracleConstantModule.ORADYN_READONLY)
#End If
    ABITKB = Nz(dyn.Fields("ABITKB"))
    Call dyn.Close
End Property

Public Property Get sysDate(Optional ByVal vFormat As String = "YYYY-MM-DD HH24:MI:SS") As String
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    sql = "SELECT "
    sql = sql & " TO_CHAR(SYSDATE,'" & vFormat & "') AS GetDATE"
    sql = sql & " FROM DUAL"
#If ORA_DEBUG = 1 Then
    Set dyn = mOraDBS.CreateDynaset(sql, Options:=dynOption.ORADYN_READONLY)
#Else
    Set dyn = mOraDBS.CreateDynaset(sql, Options:=OracleConstantModule.ORADYN_READONLY)
#End If
    sysDate = dyn.Fields("GetDATE")
    Call dyn.Close
End Property

Public Property Get SystemKey() As String
    SystemKey = "SYSTEM"
End Property

Public Property Get DatabaseName() As Variant
    DatabaseName = mReg.DbDatabaseName
End Property

#If 0 Then
'//Recordset.Fields(0).FieldSize プロパティが無いので使用不可
Public Sub MaxLength(ByVal vFrm As Form, ByVal vData As ORADCLib.ORADC)
    On Error Resume Next    '漢字フィールド名認識が出来ない為
    Dim obj As Object
    For Each obj In vFrm.Controls
        If TypeOf obj Is TextBox Then
            If obj.DataField <> "" Then
                obj.MaxLength = vData.Recordset.Fields(obj.DataField).Size
            End If
        End If
    Next obj
End Sub
#End If

#If ORA_DEBUG = 1 Then
    Public Property Get Database() As OraDatabase
#Else
    Public Property Get Database() As Object
#End If
    Set Database = mOraDBS
End Property

#If ORA_DEBUG = 1 Then
    Public Property Get OpenRecordset(vSQL As String, Optional vOption As Integer = dynOption.ORADYN_DEFAULT) As OraDynaset
#Else
    Public Property Get OpenRecordset(vSQL As String, Optional vOption As Integer = OracleConstantModule.ORADYN_DEFAULT) As Object
#End If
'''2002/10/04 マウスカーソルがうっとうしいので消した
'''    Dim ms As New MouseClass
'''    Call ms.Start
    Set OpenRecordset = mOraDBS.CreateDynaset(vSQL, vOption)
'''    Set ms = Nothing
End Property

Public Property Get Nz(vData As Variant, Optional ByVal vDefault As Variant = "") As Variant
    Nz = IIf(IsNull(vData) Or "" = Trim(vData), vDefault, vData)
End Property

Public Property Get Connect()
    'OraDatabase.Connect ではユーザー名のみしか返却してくれないので作成
    Connect = mReg.DbUserName & "/" & mReg.DbPassword
End Property

Private Sub Class_Initialize()
    On Error GoTo Class_InitializeError
#If ORA_DEBUG = 1 Then
    Set mOraSSN = New OraSessionClass
    ''システムDBオープン
    Set mOraDBS = mOraSSN.OpenDatabase(mReg.DbDatabaseName, mReg.DbUserName & "/" & mReg.DbPassword, OracleInProcServer.ORADB_DEFAULT)
#Else
    Set mOraSSN = CreateObject("OracleInProcServer.XOraSession")
    ''システムDBオープン
    Set mOraDBS = mOraSSN.OpenDatabase(mReg.DbDatabaseName, mReg.DbUserName & "/" & mReg.DbPassword, OracleConstantModule.ORADB_DEFAULT)
#End If
    mConnected = True
    Exit Sub
Class_InitializeError:
    'ｴﾗｰ 3024 が示すﾃﾞｰﾀﾍﾞｰｽが見つからないようなｴﾗｰの場合は
    '高度なｴﾗｰ処理を必要とします。
    Call MsgBox(Error & "(" & Err & ")", vbCritical, "DatabaseClass.Class_Initialize()")
    '//////////////////////////////////////////////////////////
    '// システムマスタがオープン出来ないのでこれ以上処理不可能 !!
    End
    '//////////////////////////////////////////////////////////
End Sub

Private Sub Class_Terminate()
    mConnected = False
    Set mOraDBS = Nothing
    Set mOraSSN = Nothing
    Set mReg = Nothing
End Sub

#If 0 Then
Private Function GetFieldSize(vTable As String, vField As String) As Integer
    On Error GoTo GetFieldSizeError
'    GetFieldSize = mOraDBS.TableDefs(vTable).Fields(vField).FieldSize
GetFieldSizeError:
End Function
#End If

#If 0 Then
Public Function GetPrimaryKey(ByVal vvaTables As Variant, ByRef rstKeyFld() As String) As Boolean
    'For Access DB
    Dim idx As Indexes
    Dim ix As Integer, idxStr As String
    
    Set idx = mOraDBS.TableDefs(vvaTables).Indexes
    For ix = 0 To idx.Count - 1
        If idx(ix).Primary = True Then
            Exit For
        End If
    Next ix
    If ix > idx.Count - 1 Then
        Exit Function
    End If
    'Index(?).Fields => "+Key1;+Key2;+Key3"
    idxStr = Mid(idx(ix).Fields, 2)
    ix = 0
    Do While 0 <> InStr(idxStr, ";+")
        ReDim Preserve rstKeyFld(0 To ix) As String
        rstKeyFld(ix) = Mid(idxStr, 1, InStr(idxStr, ";+") - 1)
        idxStr = Mid(idxStr, InStr(idxStr, ";+") + 2)
        ix = ix + 1
    Loop
    ReDim Preserve rstKeyFld(0 To ix) As String
    rstKeyFld(ix) = idxStr
    GetPrimaryKey = True
End Function
#End If

Public Sub SetDBCombo(vobData As data, robDbCmb As DBCombo, Optional vField As String = "")
    On Error Resume Next
    If vField = "" And robDbCmb.DataField <> "" Then
        vField = robDbCmb.DataField
    End If
    robDbCmb.Text = StrConv(LeftB(StrConv(robDbCmb.Text, vbFromUnicode), vobData.Recordset.Fields(vField).Size), vbUnicode)
    On Error GoTo 0
End Sub

Public Sub BankDbListRefresh(ByVal vData As ORADC, ByVal vYomi As ComboBox, ByVal vList As DBList, ByVal vRecordKubun As Integer, Optional ByVal vSelected As String = "")
    Const cCode As Integer = 0
    Const cKanji As Integer = 1
    Const cKana As Integer = 2
    Dim sql As String
    Dim ms As New MouseClass
    Dim vName As Variant
    Call ms.Start
    
    If vRecordKubun = eBankRecordKubun.Bank Then
        vName = Array("DABANK", "DAKJNM", "daknnm")
    Else
        vName = Array("DASITN", "DAKJNM", "daknnm")
    End If
    sql = "SELECT " & vbCrLf
    sql = sql & vName(cCode) & " || ' ' || " & vName(cKanji) & " AS NameList," & vbCrLf
    sql = sql & vName(cCode) & "," & vbCrLf
    sql = sql & vName(cKana) & vbCrLf
    sql = sql & " FROM tdBankMaster" & vbCrLf
    sql = sql & " WHERE DARKBN = '" & vRecordKubun & "'" & vbCrLf
    If vYomi.Text <> "" Then
        sql = sql & " AND (" & vbCrLf & pKanaGroup(CStr(vName(cKana)), vYomi) & ")" & vbCrLf
    End If
    If vRecordKubun = eBankRecordKubun.Shiten And "" <> vSelected Then
        sql = sql & " AND DABANK = '" & vSelected & "'" & vbCrLf
    End If
'''2002/10/09 ホストデータの関係でフィールドを削除した
'''    sql = sql & "   AND TO_CHAR(SYSDATE,'YYYYMMDD') BETWEEN DAYKST AND DAYKED"  '//有効データ絞込み
    sql = sql & " GROUP BY " & vName(cCode) & " || ' ' || " & vName(cKanji) & "," & vName(cCode) & "," & vName(cKana) & vbCrLf
    '//True=カナ順 / False=コード順
    If True = mReg.BankSortOption Then
        sql = sql & " ORDER BY " & vName(cKana) & "," & vName(cCode) & vbCrLf
    Else
        sql = sql & " ORDER BY " & vName(cCode) & "," & vName(cKana) & vbCrLf
    End If
    vData.RecordSource = sql
    vData.Refresh
    vList.ListField = "NameList"
End Sub

Private Function pKanaGroup(vFieldName As String, vCombo As ComboBox) As String
    Dim sql As String
    
    Dim mKana As String, i As Integer
    '//実際のアスキーコード順
    'ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ
    mKana = "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖ@@ﾗﾘﾙﾚﾛﾜｦﾝ"  'ﾔﾕﾖ + @@ は 5 ステップにするため作成
    
    If vCombo.ListIndex < vCombo.ListCount Then
        For i = (vCombo.ListIndex - 1) * 5 To (vCombo.ListIndex) * 5 - 1
'//2006/04/24 ワ行を選択時全件対象となるバグを修正＆ヤユヨ＠＠の「＠」も省く
            If "" <> Mid(mKana, i + 1, 1) And "@" <> Mid(mKana, i + 1, 1) Then
                sql = sql & " " & vFieldName & " LIKE '" & Mid(mKana, i + 1, 1) & "%'" & vbCrLf & " OR"
            End If
        Next i
        sql = Left(sql, Len(sql) - Len("OR"))
    Else
        sql = " 1 = 1 " & vbCrLf        ' 1=1 => 無条件ヒット
    End If
    pKanaGroup = sql
End Function

#If ORA_DEBUG = 1 Then
    Public Function ErrorCheck(Optional ByVal vDb As OraDatabase = Nothing) As Boolean
#Else
    Public Function ErrorCheck(Optional ByVal vDb As Object = Nothing) As Boolean
#End If
    If TypeName(vDb) = "Nothing" Then
        If mOraDBS.LastServerErr Then
            Call Me.MsgBox("内部エラーが発生しました.(Error Code = " & mOraDBS.LastServerErr & ")" & vbCrLf & mOraDBS.LastServerErrText & vbCrLf & vbCrLf & cSystemMessage, _
                        vbCritical + vbOKOnly, mReg.Title)
            ErrorCheck = True
        ElseIf Err.Number Then
            Call Me.MsgBox("内部エラーが発生しました.(Error Code = " & Err.Number & ")" & vbCrLf & Err.Description & vbCrLf & vbCrLf & cSystemMessage, _
                        vbCritical + vbOKOnly, mReg.Title)
            ErrorCheck = True
        End If
    ElseIf vDb.LastServerErr Then
        Select Case vDb.LastServerErr
        Case 54
            Call Me.MsgBox("既に他のユーザーが使用しています.(Error Code = " & vDb.LastServerErr & ")" & vbCrLf & vDb.LastServerErrText & vbCrLf & vbCrLf & cSystemMessage, _
                        vbCritical + vbOKOnly, mReg.Title)
        Case Else
            Call Me.MsgBox("内部エラーが発生しました.(Error Code = " & vDb.LastServerErr & ")" & vbCrLf & vDb.LastServerErrText & vbCrLf & vbCrLf & cSystemMessage, _
                        vbCritical + vbOKOnly, mReg.Title)
        End Select
        ErrorCheck = True
    ElseIf Err.Number Then
        Call Me.MsgBox("内部エラーが発生しました.(Error Code = " & Err.Number & ")" & vbCrLf & Err.Description & vbCrLf & vbCrLf & cSystemMessage, _
                    vbCritical + vbOKOnly, mReg.Title)
        ErrorCheck = True
    End If
    Call mOraDBS.Rollback
End Function

Public Function FirstDay(ByVal vDate As Long) As Long
    If vDate > 19000101 Then
        FirstDay = (Int(vDate / 100) * 100) + 1
    End If
End Function

Public Function LastDay(ByVal vDate As Long, Optional vAddMonths As Integer = 0) As Long
    If vDate > 19000101 Then
#If ORA_DEBUG = 1 Then
        Dim dyn As OraDynaset
        Set dyn = mOraDBS.CreateDynaset("SELECT TO_CHAR(LAST_DAY(ADD_MONTHS(TO_DATE(" & FirstDay(vDate) & ",'YYYYMMDD')," & vAddMonths & ")),'YYYYMMDD') AS LastDay FROM DUAL", dynOption.ORADYN_READONLY)
#Else
        Dim dyn As Object
        Set dyn = mOraDBS.CreateDynaset("SELECT TO_CHAR(LAST_DAY(ADD_MONTHS(TO_DATE(" & FirstDay(vDate) & ",'YYYYMMDD')," & vAddMonths & ")),'YYYYMMDD') AS LastDay FROM DUAL", OracleConstantModule.ORADYN_READONLY)
#End If
        LastDay = dyn.Fields("LastDay")
        Set dyn = Nothing
    Else
        LastDay = 20991231
    End If
End Function

Public Property Get FieldNames(ByVal vTable As String, Optional vAddConditions As String = "") As Variant
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim columns() As Variant, cnt As Integer
    sql = "SELECT COLUMN_NAME FROM USER_TAB_COLUMNS"
    sql = sql & " WHERE TABLE_NAME = '" & UCase(vTable) & "'"
    If "" <> vAddConditions Then
        sql = sql & " " & vAddConditions
    End If
    sql = sql & " ORDER BY COLUMN_ID"
#If ORA_DEBUG = 1 Then
    Set dyn = OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    Erase columns
    Do Until dyn.EOF()
        ReDim Preserve columns(cnt) As Variant
        columns(cnt) = dyn.Fields("COLUMN_NAME")
        cnt = cnt + 1
        Call dyn.MoveNext
    Loop
    Set dyn = Nothing
    FieldNames = columns
End Property

Public Property Get SystemUpdate(ByVal vField As String) As Variant
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    sql = "SELECT * FROM taSystemInformation"
    sql = sql & " WHERE AASKEY = '" & Me.SystemKey & "'"
#If ORA_DEBUG = 1 Then
    Set dyn = Me.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = Me.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If Not dyn.EOF() Then
        SystemUpdate = dyn.Fields(vField)
    End If
    Call dyn.Close
    Set dyn = Nothing
End Property

Public Property Let SystemUpdate(ByVal vField As String, ByVal vNewValue As Variant)
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    sql = "UPDATE taSystemInformation SET "
    Select Case UCase(vField)
    Case UCase("aaupdE")
        sql = sql & " AAUPD1 = 0,"
        sql = sql & " AAUPD2 = 0,"
        sql = sql & " AAUPD3 = 0,"
'//2004/04/13 累積日を次回「新規データ」扱い日に設定
'//        sql = sql & " AANWDT = AAYTDT,"     '//口座振替予定 作成日を新規データの判断基準日に更新
        sql = sql & " AANWDT = SYSDATE,"
    Case UCase("AAUPD1")
'//2003/02/03 口座振替予定 作成日を更新
        sql = sql & " AAYTDT = SYSDATE,"
    Case Else
    End Select
    sql = sql & vField & " = '" & vNewValue & "',"
    sql = sql & "AAUSID = '" & LoginUserName & "',"
    sql = sql & "AAUPDT = SYSDATE"
    sql = sql & " WHERE AASKEY = '" & Me.SystemKey & "'"
    Call mOraDBS.ExecuteSQL(sql)
End Property

Public Property Get CheckDateType(ByVal vData As Variant) As String
    On Error GoTo CheckDateTypeError
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset, tmp As String
#Else
    Dim sql As String, dyn As Object, tmp As String
#End If
'//////////////////////////////////////////////////////////////////////////////////////////
'//注意：コントロールパネル/地域/日付の短い形式のプロパティがに従ってフォーマットされる.
'//////////////////////////////////////////////////////////////////////////////////////////
    '//一旦この形式で文字列にしないと「０３年」が本来２００３年となる所がオラクルでは ->０００３年となってしまう.
    '//先に日付形式でフォーマットしないと SQL 文のエラーとなって例外が起きる.
    tmp = Format(CVDate(vData), "yyyy/mm/dd hh:nn:ss")
    sql = "SELECT TO_CHAR(TO_DATE('" & tmp & "','YY/MM/DD HH24:MI:SS'),'YYYY/MM/DD HH24:MI:SS') AS CheckDate"
    sql = sql & " FROM DUAL"
#If ORA_DEBUG = 1 Then
    Set dyn = Me.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = Me.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If Not dyn.EOF() Then
        CheckDateType = dyn.Fields("CheckDate")
    End If
    Call dyn.Close
    Set dyn = Nothing
    Exit Property
CheckDateTypeError:
    CheckDateType = ""
End Property

'//////////////////////////////////////////////////////////////////////////
'//ＳＱＬ生成時のみ使用してください。：ＶＢ条件判断には使用しないで下さい。
Public Function ColumnDataSet( _
            vData As Variant, _
            Optional vType As String = "S", _
            Optional vEnd As Boolean = False _
        ) As Variant
    Dim vTemp As Variant
    If Not IsNull(vData) Then
        vTemp = Trim(Replace(vData, Chr(0), Chr(32)))
    End If
    If "" = Trim(vTemp) Or IsEmpty(vTemp) Then
        ColumnDataSet = "NULL"
    Else
        Select Case UCase(vType)
        Case "S"        '//文字
            ColumnDataSet = "'" & Trim(vTemp) & "'"
        Case "I", "L"   '//数値
            ColumnDataSet = Trim(vTemp)
        Case "D"        '//日付
            ColumnDataSet = "'" & Format(Trim(vTemp), "yyyy/MM/dd hh:nn:ss") & "'"
        End Select
    End If
    If False = vEnd Then
        ColumnDataSet = ColumnDataSet & ","
    End If
End Function

Public Sub TriggerControl(vTable As String, Optional vEnable As Boolean = True)
'//2007/06/18 全てログを取りたいのでコメント化！
#If 0 Then
'''#If ORA_DEBUG = 1 Then
'''    Dim sql As String, dyn As OraDynaset
'''#Else
'''    Dim sql As String, dyn As Object
'''#End If
'''    Dim mode As String
'''    If True = vEnable Then
'''        mode = " ENABLE"
'''    Else
'''        mode = " DISABLE"
'''    End If
'''    sql = "SELECT TRIGGER_NAME FROM USER_TRIGGERS "
'''    sql = sql & " WHERE TABLE_NAME = '" & UCase(vTable) & "'"
'''#If ORA_DEBUG = 1 Then
'''    Set dyn = Me.OpenRecordset(sql, dynOption.ORADYN_READONLY)
'''#Else
'''    Set dyn = Me.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
'''#End If
'''    Do Until dyn.EOF
'''        Call Me.Database.ExecuteSQL("ALTER TRIGGER " & dyn.Fields("TRIGGER_NAME").Value & mode)
'''        Call dyn.MoveNext
'''    Loop
'''    Call dyn.Close
#End If
End Sub


