VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "バージョン情報"
   ClientHeight    =   3045
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5160
   ClipControls    =   0   'False
   Icon            =   "ABOUT.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2101.713
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   4845.507
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   180
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "ABOUT.frx":0442
      ScaleHeight     =   337.12
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3600
      TabIndex        =   0
      Top             =   2085
      Width           =   1365
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "ｼｽﾃﾑ情報(&S)..."
      Height          =   345
      Left            =   3600
      TabIndex        =   2
      Top             =   2535
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '実線 (ふちどり)
      Index           =   1
      X1              =   84.515
      X2              =   4749.724
      Y1              =   1325.218
      Y2              =   1325.218
   End
   Begin VB.Label lblFileDescription 
      Caption         =   "アプリケーションの説明"
      ForeColor       =   &H00000000&
      Height          =   990
      Left            =   1050
      TabIndex        =   3
      Top             =   825
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "アプリケーション タイトル"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   5
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "バージョン"
      Height          =   225
      Left            =   1050
      TabIndex        =   6
      Top             =   540
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "警告. ....."
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   4
      Top             =   2085
      Width           =   3270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   84.515
      X2              =   4750.663
      Y1              =   1325.218
      Y2              =   1325.218
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "編集"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "コピー(&C)"
      End
      Begin VB.Menu mnuAllSelect 
         Caption         =   "すべて選択(&A)"
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ﾚｼﾞｽﾄﾘ ｷｰ ｾｷｭﾘﾃｨ ｵﾌﾟｼｮﾝ...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ﾚｼﾞｽﾄﾘ ｷｰ ROOT 型...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode Null 文字で終わる文字列
Const REG_DWORD = 4                      ' 32 ﾋﾞｯﾄ数値

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

#If 0 Then
Public mvaSS As vaSpread
#End If

Private Sub cmdSysInfo_Click()
    On Error GoTo cmdSysInfo_ClickError
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' ﾚｼﾞｽﾄﾘからｼｽﾃﾑ情報ﾌﾟﾛｸﾞﾗﾑのﾊﾟｽ\名前を取得します...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' ﾚｼﾞｽﾄﾘからｼｽﾃﾑ情報ﾌﾟﾛｸﾞﾗﾑのﾊﾟｽ名のみを取得します...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' 既に存在するはずの 32 ﾋﾞｯﾄ ﾊﾞｰｼﾞｮﾝのﾌｧｲﾙを確認します。
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"        ' ｴﾗｰ - ﾌｧｲﾙが見つかりません...
        Else
            GoTo cmdSysInfo_ClickError
        End If                                                 ' ｴﾗｰ - ﾚｼﾞｽﾄﾘ ｴﾝﾄﾘが見つかりません...
    Else
        GoTo cmdSysInfo_ClickError
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
cmdSysInfo_ClickError:
    Call gdDBS.MsgBox("現時点ではｼｽﾃﾑ情報を使用できません", vbOKOnly, Me.Caption)
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim reg As New RegistryClass
    cmnDlg.HelpFile = App.HelpFile
    cmnDlg.HelpCommand = &HB  'cdlHelpKey  'cdlHelpContext
    
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Me.Caption = App.Comments & "のバージョン情報"
    lblTitle.Caption = App.Comments
    lblVersion.Caption = "バージョン " & App.Major & "." & Format(App.Minor, "00.") & Format(App.Revision, "00")
    lblFileDescription.Caption = reg.CompanyName & vbCrLf & _
            UCase(App.EXEName) & ".EXE (" & FileDateTime(App.path & "\" & App.EXEName & ".EXE") & ")" '&
    lblFileDescription.Caption = lblFileDescription.Caption & vbCrLf & "[ User = " & reg.DbUserName & " ] [ Connect = " & reg.DbDatabaseName & " ]"
    If reg.Debuged = True Then
        lblFileDescription.Caption = lblFileDescription.Caption & vbCrLf & "Debug Mode=True"
    End If
    lblDisclaimer = lblDisclaimer & vbCrLf & vbCrLf & "ﾀﾞｲﾔﾓﾝﾄﾞｺﾝﾋﾟｭｰﾀｰｻｰﾋﾞｽ(株)大阪支店" & vbCrLf & " TEL 06-6454-5977 FAX 06-6454-5990"
End Sub

Public Sub StartSysInfo()
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' ﾙｰﾌﾟ ｶｳﾝﾀ
    Dim rc As Long                                          ' 戻り値のｺｰﾄﾞ
    Dim hKey As Long                                        ' ｵｰﾌﾟﾝしたﾚｼﾞｽﾄﾘ ｷｰのﾊﾝﾄﾞﾙ
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' ﾚｼﾞｽﾄﾘ ｷｰのﾃﾞｰﾀ型
    Dim tmpVal As String                                    ' ﾚｼﾞｽﾄﾘ ｷｰ値の一時保存領域
    Dim KeyValSize As Long                                  ' ﾚｼﾞｽﾄﾘ ｷｰ変数のｻｲｽﾞ
    '------------------------------------------------------------
    ' ﾙｰﾄ ｷｰ {HKEY_LOCAL_MACHINE...} にあるﾚｼﾞｽﾄﾘ ｷｰを開きます。
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ﾚｼﾞｽﾄﾘ ｷｰを開く
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ﾊﾝﾄﾞﾙ ｴﾗｰ...
    
    tmpVal = String$(1024, 0)                             ' 変数領域の割り当て
    KeyValSize = 1024                                       ' 変数のｻｲｽﾞを記憶
    
    '------------------------------------------------------------
    ' ﾚｼﾞｽﾄﾘ ｷｰ値を取得します...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' ｷｰ値の取得/作成
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ﾊﾝﾄﾞﾙ ｴﾗｰ
    
    tmpVal = Left(tmpVal, InStr(tmpVal, Chr(0)) - 1)
    '------------------------------------------------------------
    ' 変換のために、ｷｰ値の型を調べます...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' ﾃﾞｰﾀ型検索...
    Case REG_SZ                                             ' String ﾚｼﾞｽﾄﾘ ｷｰ ﾃﾞｰﾀ型
        KeyVal = tmpVal                                     ' String 値をｺﾋﾟｰ
    Case REG_DWORD                                          ' Double Word ﾚｼﾞｽﾄﾘ ｷｰ ﾃﾞｰﾀ型
        For i = Len(tmpVal) To 1 Step -1                    ' 各ﾋﾞｯﾄの変換
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Char ごとに値を作成
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Double Word を String に変換
    End Select
    
    GetKeyValue = True                                      ' 正常終了
    rc = RegCloseKey(hKey)                                  ' ﾚｼﾞｽﾄﾘ ｷｰをｸﾛｰｽﾞ
    Exit Function                                           ' 終了
    
GetKeyError:      ' ｴﾗｰ発生後の後始末...
    KeyVal = ""                                             ' 戻り値の値を空文字列に設定
    GetKeyValue = False                                     ' 異常終了
    rc = RegCloseKey(hKey)                                  ' ﾚｼﾞｽﾄﾘ ｷｰをｸﾛｰｽﾞ
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = (UnloadMode = vbFormControlMenu)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAbout = Nothing
End Sub

#If 0 Then
'/////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////
'//各フォーム上の Spread データコピー用のポップアップメニュー
Private Sub mnuAllSelect_Click()
    With mvaSS
        '// 行 / 列 見出しの有無を判断
        .Col = IIf(.DisplayRowHeaders = True, -1, 1)
        .Row = IIf(.DisplayColHeaders = True, -1, 1)
        .Col2 = IIf(.DisplayRowHeaders = True, -1, .MaxCols)
        .Row2 = IIf(.DisplayColHeaders = True, -1, .MaxRows)
        .Action = 2 'SS_ACTION_SELECT_BLOCK
    End With
End Sub

Private Sub mnuCopy_Click()
    mvaSS.Action = 22 'SS_ACTION_CLIPBOARD_COPY
End Sub
'//各フォーム上の Spread データコピー用のポップアップメニュー
'/////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////
#End If

