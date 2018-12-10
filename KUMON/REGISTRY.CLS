VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Get CompanyName() As Variant
    CompanyName = GetSetting(App.Title, "System", "CompanyName", "")
    If Trim(CompanyName) = "" Then
        CompanyName = "No CompanyName"
        Call SaveSetting(App.Title, "System", "CompanyName", CompanyName)
    End If
End Property

Public Property Get Title() As Variant
    Title = GetSetting(App.Title, "System", "Title", "")
    If Trim(Title) = "" Then
        Title = App.Title
        Call SaveSetting(App.Title, "System", "Title", Title)
    End If
End Property

Public Property Get DbDatabaseName() As Variant
    DbDatabaseName = GetSetting(App.Title, "System", "DbDatabaseName", "")
    If DbDatabaseName = "" Then
        DbDatabaseName = "O8WS32i"
        Call SaveSetting(App.Title, "System", "DbDatabaseName", DbDatabaseName)
    End If
End Property

Public Property Get DbUserName() As Variant
    DbUserName = GetSetting(App.Title, "System", "DbUserName", "")
    If DbUserName = "" Then
        DbUserName = "admin"
        Call SaveSetting(App.Title, "System", "DbUserName", DbUserName)
    End If
End Property

Public Property Get DbPassword() As Variant
    DbPassword = GetSetting(App.Title, "System", "DbPassword", "")
    If DbPassword = "" Then
        DbPassword = "admin"
        Call SaveSetting(App.Title, "System", "DbPassword", DbPassword)
    End If
End Property

Public Property Get WarnningCount() As Variant
    WarnningCount = GetSetting(App.Title, "System", "過剰印刷警告件数", "")
    If Trim(WarnningCount) = "" Then
        WarnningCount = 10
        Call SaveSetting(App.Title, "System", "過剰印刷警告件数", WarnningCount)
    End If
End Property

Public Property Let WarnningCount(vvaData As Variant)
    Call SaveSetting(App.Title, "System", "過剰印刷警告件数", vvaData)
End Property

Public Function GetPrintMargin(vKey As Variant, vName As Variant, DefaultValue As Variant) As Variant
    '// vKey  = フォーム名称
    '// vName = "Left" Or "Top"
    '// DefaultValue = デフォルト値
    GetPrintMargin = GetSetting(App.Title, "System\" & vKey, vName & "Margin", DefaultValue)
End Function

Public Sub SetPrintMargin(vKey As Variant, vName As Variant, vValue As Variant)
    '// vKey  = フォーム名称
    '// vName = "Left" Or "Top"
    '// vValue = 設定値
    Call SaveSetting(App.Title, "System\" & vKey, vName & "Margin", vValue)
End Sub

#If 0 Then
Public Sub GetColumns(vSS As vaSpread)
    Dim Width As String, Col As Long
    Width = GetSetting(App.Title, "System", vSS.Tag, "")
    If Trim(Width) = "" Then
        Exit Sub
    End If
    For Col = 1 To vSS.MaxCols
        vSS.ColWidth(Col) = Val(Width)
        Width = Mid(Width, InStr(Width, ",") + 1)
    Next Col
End Sub
#End If

#If 0 Then
Public Sub SetColumns(vSS As vaSpread)
    Dim Width As String, Col As Long
    For Col = 1 To vSS.MaxCols
        Width = Width & vSS.ColWidth(Col) & ","
    Next Col
    Width = Left(Width, Len(Width) - 1)
    Call SaveSetting(App.Title, "System", vSS.Tag, Width)
End Sub
#End If

Public Property Get FurikaeDataImport() As Variant
    FurikaeDataImport = GetSetting(App.Title, "System", "前回取得振替日", "")
    If FurikaeDataImport = "" Then
        FurikaeDataImport = Format(Now(), "YYYY/MM/DD")
        Call SaveSetting(App.Title, "System", "前回取得振替日", FurikaeDataImport)
    End If
End Property

Public Property Let FurikaeDataImport(vNewData As Variant)
    Call SaveSetting(App.Title, "System", "前回取得振替日", FurikaeDataImport)
End Property

Public Function GetFormPosition(ByVal vCaption As String) As String
    GetFormPosition = GetSetting(App.Title, "FormPositions", vCaption, "")
End Function

Public Sub SetFormPosition(ByVal vCaption As String, ByVal vNewValue As Variant)
    Call SaveSetting(App.Title, "FormPositions", vCaption, vNewValue)
End Sub

Public Property Get InputFileName(ByVal vSection As String) As Variant
    InputFileName = GetSetting(App.Title, "InputFileName", vSection, vSection & ".txt")
End Property

Public Property Let InputFileName(ByVal vSection As String, ByVal vNewValue As Variant)
    Call SaveSetting(App.Title, "InputFileName", vSection, vNewValue)
End Property

Public Property Get OutputFileName(ByVal vSection As String) As Variant
    OutputFileName = GetSetting(App.Title, "OutputFileName", vSection, vSection & ".txt")
End Property

Public Property Let OutputFileName(ByVal vSection As String, ByVal vNewValue As Variant)
    Call SaveSetting(App.Title, "OutputFileName", vSection, vNewValue)
End Property

Public Property Get TransferCommand(ByVal vSection As String) As Variant
    TransferCommand = GetSetting(App.Title, "TransferCommand", vSection, "")
    If "" = TransferCommand Then
        TransferCommand = "C:\Program Files\Internet Explorer\IEXPLORE.EXE"
        Call SaveSetting(App.Title, "TransferCommand", vSection, TransferCommand)
    End If
End Property

'//2006/03/02 メニューをタブ形式にしたので表示タブ位置を設定できるように
Public Property Get MenuTab() As Variant
    MenuTab = GetSetting(App.Title, "System", "MenuTab", "0")
End Property

Public Property Let MenuTab(ByVal vNewValue As Variant)
    Call SaveSetting(App.Title, "System", "MenuTab", vNewValue)
End Property

Public Property Get MenuButton() As Variant
    MenuButton = GetSetting(App.Title, "System", "MenuButton", "0")
End Property

Public Property Let MenuButton(ByVal vNewValue As Variant)
    Call SaveSetting(App.Title, "System", "MenuButton", vNewValue)
End Property

Public Property Get Debuged() As Variant
    Debuged = GetSetting(App.Title, "System", "Debug", "")
    If "" = Debuged Then
        Debuged = "False"
        Call SaveSetting(App.Title, "System", "Debug", Debuged)
    End If
End Property

Public Property Get zaAutologOut() As Variant
    zaAutologOut = GetSetting(App.Title, "System", "zaAutologOut", "NonRegistry")
    If "NonRegistry" = zaAutologOut Then
        zaAutologOut = "True"
        Call SaveSetting(App.Title, "System", "zaAutologOut", zaAutologOut)
    End If
    zaAutologOut = (zaAutologOut = True)
End Property

Public Property Get BankSortOption() As Variant
    BankSortOption = GetSetting(App.Title, "System", "BankSortOption", "")
    If "" = BankSortOption Then
        BankSortOption = "True"
        Call SaveSetting(App.Title, "System", "BankSortOption", BankSortOption)
    End If
End Property

Public Property Get LzhExtractFile() As Variant
    LzhExtractFile = GetSetting(App.Title, "System", "LzhExtractFile", "")
    If "" = LzhExtractFile Then
        LzhExtractFile = "C:\"
        Call SaveSetting(App.Title, "System", "LzhExtractFile", LzhExtractFile)
    End If
End Property

Public Property Get CheckTimer() As Integer
    CheckTimer = GetSetting(App.Title, "System", "CheckTimer", 0)
    If 0 = CheckTimer Then
        CheckTimer = 10     '// １０分
        Call SaveSetting(App.Title, "System", "CheckTimer", CheckTimer)
    End If
End Property

