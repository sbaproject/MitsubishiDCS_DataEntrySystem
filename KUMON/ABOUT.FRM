VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�o�[�W�������"
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
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   4845.507
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
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
      ScaleMode       =   0  'հ�ް
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
      Caption         =   "���я��(&S)..."
      Height          =   345
      Left            =   3600
      TabIndex        =   2
      Top             =   2535
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '���� (�ӂ��ǂ�)
      Index           =   1
      X1              =   84.515
      X2              =   4749.724
      Y1              =   1325.218
      Y2              =   1325.218
   End
   Begin VB.Label lblFileDescription 
      Caption         =   "�A�v���P�[�V�����̐���"
      ForeColor       =   &H00000000&
      Height          =   990
      Left            =   1050
      TabIndex        =   3
      Top             =   825
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "�A�v���P�[�V���� �^�C�g��"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   5
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "�o�[�W����"
      Height          =   225
      Left            =   1050
      TabIndex        =   6
      Top             =   540
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "�x��. ....."
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
      Caption         =   "�ҏW"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "�R�s�[(&C)"
      End
      Begin VB.Menu mnuAllSelect 
         Caption         =   "���ׂđI��(&A)"
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ڼ޽�� �� ����è ��߼��...
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
                     
' ڼ޽�� �� ROOT �^...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode Null �����ŏI��镶����
Const REG_DWORD = 4                      ' 32 �ޯĐ��l

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
    
    ' ڼ޽�؂��缽�я����۸��т��߽\���O���擾���܂�...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' ڼ޽�؂��缽�я����۸��т��߽���݂̂��擾���܂�...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' ���ɑ��݂���͂��� 32 �ޯ� �ް�ޮ݂�̧�ق��m�F���܂��B
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"        ' �װ - ̧�ق�������܂���...
        Else
            GoTo cmdSysInfo_ClickError
        End If                                                 ' �װ - ڼ޽�� ���؂�������܂���...
    Else
        GoTo cmdSysInfo_ClickError
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
cmdSysInfo_ClickError:
    Call gdDBS.MsgBox("�����_�łͼ��я����g�p�ł��܂���", vbOKOnly, Me.Caption)
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
    Me.Caption = App.Comments & "�̃o�[�W�������"
    lblTitle.Caption = App.Comments
    lblVersion.Caption = "�o�[�W���� " & App.Major & "." & Format(App.Minor, "00.") & Format(App.Revision, "00")
    lblFileDescription.Caption = reg.CompanyName & vbCrLf & _
            UCase(App.EXEName) & ".EXE (" & FileDateTime(App.path & "\" & App.EXEName & ".EXE") & ")" & vbCrLf & _
             "Database = " & reg.DbDatabaseName & vbCrLf & "UserName = " & reg.DbUserName & vbCrLf & _
             "LoginUser = " & gdDBS.LoginUserName
    If reg.Debuged = True Then
        lblFileDescription.Caption = lblFileDescription.Caption & vbCrLf & "Debug Mode=True"
    End If
    lblDisclaimer = lblDisclaimer & vbCrLf & vbCrLf & "�޲����޺��߭������޽(��)���x�X" & vbCrLf & " TEL 06-6454-5977 FAX 06-6454-5990"
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' ٰ�� ����
    Dim rc As Long                                          ' �߂�l�̺���
    Dim hKey As Long                                        ' ����݂���ڼ޽�� ���������
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' ڼ޽�� �����ް��^
    Dim tmpVal As String                                    ' ڼ޽�� ���l�̈ꎞ�ۑ��̈�
    Dim KeyValSize As Long                                  ' ڼ޽�� ���ϐ��̻���
    '------------------------------------------------------------
    ' ٰ� �� {HKEY_LOCAL_MACHINE...} �ɂ���ڼ޽�� �����J���܂��B
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ڼ޽�� �����J��
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ����� �װ...
    
    tmpVal = String$(1024, 0)                             ' �ϐ��̈�̊��蓖��
    KeyValSize = 1024                                       ' �ϐ��̻��ނ��L��
    
    '------------------------------------------------------------
    ' ڼ޽�� ���l���擾���܂�...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' ���l�̎擾/�쐬
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ����� �װ
    
    tmpVal = Left(tmpVal, InStr(tmpVal, Chr(0)) - 1)
    '------------------------------------------------------------
    ' �ϊ��̂��߂ɁA���l�̌^�𒲂ׂ܂�...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' �ް��^����...
    Case REG_SZ                                             ' String ڼ޽�� �� �ް��^
        KeyVal = tmpVal                                     ' String �l���߰
    Case REG_DWORD                                          ' Double Word ڼ޽�� �� �ް��^
        For i = Len(tmpVal) To 1 Step -1                    ' �e�ޯĂ̕ϊ�
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Char ���Ƃɒl���쐬
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Double Word �� String �ɕϊ�
    End Select
    
    GetKeyValue = True                                      ' ����I��
    rc = RegCloseKey(hKey)                                  ' ڼ޽�� ����۰��
    Exit Function                                           ' �I��
    
GetKeyError:      ' �װ������̌�n��...
    KeyVal = ""                                             ' �߂�l�̒l���󕶎���ɐݒ�
    GetKeyValue = False                                     ' �ُ�I��
    rc = RegCloseKey(hKey)                                  ' ڼ޽�� ����۰��
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
'//�e�t�H�[����� Spread �f�[�^�R�s�[�p�̃|�b�v�A�b�v���j���[
Private Sub mnuAllSelect_Click()
    With mvaSS
        '// �s / �� ���o���̗L���𔻒f
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
'//�e�t�H�[����� Spread �f�[�^�R�s�[�p�̃|�b�v�A�b�v���j���[
'/////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////
#End If
