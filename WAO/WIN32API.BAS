Attribute VB_Name = "Win32API_Module"
Option Explicit

Declare Sub PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal msg_id As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
'Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Public Const HWND_TOPMOST = (-1)
'Public Const SWP_NOSIZE = &H1&
'Public Const SWP_NOMOVE = &H2&

''// IME �֘A API
Declare Function IMMGetContext Lib "imm32.dll" Alias "ImmGetContext" (ByVal hWnd As Long) As Long
Declare Function IMMReleaseContext Lib "imm32.dll" Alias "ImmReleaseContext" (ByVal hWnd As Long, ByVal hImc As Long) As Long
Declare Function IMMSetOpenStatus Lib "imm32.dll" Alias "ImmSetOpenStatus" (ByVal hImc As Long, ByVal b As Long) As Long

'// �V�X�e����� �֘A
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Const SM_CYMENU = 15     '���j���[�̍��� �擾�萔
Public Const SM_CYCAPTION = 4   '�^�C�g���̍��� �擾�萔
Public Const SM_CYBORDER = 6    '�g�̍���       �擾�萔
Public Type SYSTEM_METRICS
        Border  As Long     '�E�C���h�E�̘g
        Caption As Long     '�^�C�g���̍���
        Menu    As Long     '���j���[�̍���
        TaskBar As Long     '�^�X�N�o�[�̍���
End Type

Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As String

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
        ByVal lpBuffer As String, _
        nSize As Long _
    ) As Long

Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
    Public Const MOVEFILE_REPLACE_EXISTING = &H1
    Public Const MOVEFILE_COPY_ALLOWED = &H2
    Public Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4
' @(f)
'
' �@�\      : FEP ����
'
' �Ԃ�l    : ����
'
' ������    : ARG1 - blnMode OFF=FALSE / ON=TRUE
'
' �@�\����  : ���{����͂̃I��/�I�t�𐧌䂷��
'
' ���l      :
'
Public Sub SwitchIME(hWnd As Long, ByVal blnMode As Boolean)
    Dim hImc As Long
    
    hImc = IMMGetContext(Screen.ActiveForm.hWnd)
    Call IMMSetOpenStatus(hImc, blnMode)
    Call IMMReleaseContext(Screen.ActiveForm.hWnd, hImc)
End Sub

' @(f)
'
' �@�\      : �w�肵���E�C���h�E�̃L���v�V�������������E�C���h�E���A�N�e�B�u�ɂ���
'
' �Ԃ�l    : TRUE = ����I��
'
' ������    : ARG1 -strCaption:��������E�C���h�E�̃L���v�V����
'
' �@�\����  : �w�肵���E�C���h�E�̃L���v�V�������������E�C���h�E���A�N�e�B�u�ɂ���
'
' ���l      :
'
Public Function SetTopWindow(ByVal strCaption As String) As Boolean
    ' �E�B���h�E���őO�ʂɐݒ�
    SetTopWindow = BringWindowToTop(GetWindowTextHandle(strCaption))
    'SetTopWindow = SetWindowPos(GetWindowTextHandle(strCaption), HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE OR SWP_NOSIZE)
End Function


' @(f)
'
' �@�\      : �w�肵���E�C���h�E�̃L���v�V�������������E�C���h�E����������
'
' �Ԃ�l    : �E�C���h�E�̃n���h��
'
' ������    : ARG1 -strCaption:��������E�C���h�E�̃L���v�V����
'
' �@�\����  : �w�肵���E�C���h�E�̃L���v�V�������������E�C���h�E��T��
'
' ���l      :
'
Public Function GetWindowTextHandle(ByVal strCaption As String) As Long
    Dim Dst  As String
    Dim Src  As String
    Dim tmp  As String * 256
    Dim hWnd As Long
    Dim Mach As Boolean
      
    Dst = UCase(strCaption)
    hWnd = FindWindow(vbNullString, vbNullString)   'vbNullString ���}���R�����X
    hWnd = GetWindow(hWnd, GW_HWNDFIRST)
    Do Until hWnd = 0
        Call GetWindowText(hWnd, tmp, Len(tmp))
        Src = UCase(Left(tmp, InStr(tmp, vbNullChar) - 1))  'NULL �ŏI����Ă��镶����Ȃ̂�...<< Left(str,�Ԃ�l�̒���)�ł͐��m�Ɏ��Ȃ� >>
    #If 0 Then
        Dim ClsStr As String * 256
        Call GetClassName(hWnd, ClsStr, Len(ClsStr))
        Debug.Print "hWnd="; hWnd; " Src=["; Trim(Src); "]"; " Class=["; Trim(Left(ClsStr, InStr(ClsStr, vbNullChar) - 1)); "]"
    #End If
        If InStr(Src, Dst) > 0 Then
            GetWindowTextHandle = hWnd
            Exit Function
        End If
        hWnd = GetNextWindow(hWnd, GW_HWNDNEXT)
    Loop
    GetWindowTextHandle = 0
End Function

#If 0 Then
Public Function TaskBarPosition(ByRef sm As SYSTEM_METRICS) As Integer
    Dim rc As RECT
    Call SystemParametersInfo(SPI_GETWORKAREA, 0, rc, 0)
    sm.Border = GetSystemMetrics(SM_CYBORDER) * 16      'Screen.TwipsPerPixelY
    sm.Caption = GetSystemMetrics(SM_CYCAPTION) * 16    'Screen.TwipsPerPixelY
    sm.Menu = GetSystemMetrics(SM_CYMENU) * 16          'Screen.TwipsPerPixelY
    '///////////////////////////////////////////////
    '// Pos  = �^�X�N�o�[���ǂ̏ꏊ�ɑ��݂��邩( 3=�� / 2=�� / 1=�E / 0=�� / -1=�������ɉB��Ă���(�ꏊ�͕�����Ȃ�)
    '// Size = �^�X�N�o�[�̕��A���͍���
    If rc.Left <> 0 Then
        '���Ƀ^�X�N�o�[������
        TaskBarPosition = 3
        sm.TaskBar = rc.Left * Screen.TwipsPerPixelX
    ElseIf rc.Top <> 0 Then
        '��Ƀ^�X�N�o�[������
        TaskBarPosition = 2
        sm.TaskBar = rc.Top * Screen.TwipsPerPixelY
    ElseIf rc.Right * Screen.TwipsPerPixelX <> Screen.Width Then
        '�E�Ƀ^�X�N�o�[������
        TaskBarPosition = 1
        sm.TaskBar = Screen.Width - rc.Right * Screen.TwipsPerPixelX
    ElseIf rc.Bottom * Screen.TwipsPerPixelY <> Screen.Height Then
        '���Ƀ^�X�N�o�[������
        TaskBarPosition = 0
        sm.TaskBar = Screen.Height - rc.Bottom * Screen.TwipsPerPixelY
    Else
        TaskBarPosition = -1
        sm.TaskBar = 0
    End If
#If 0 Then
    MsgBox "left=" & rc.Left & " top=" & rc.Top & " right=" & rc.Right & " bottom=" & rc.Bottom _
        & " Screen= H " & Screen.Height & " W " & Screen.Width
#End If
End Function
#End If
