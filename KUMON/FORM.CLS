VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mReg As New RegistryClass
'Private Const pcFormWidth As Integer = 12000   '800x600
'Private Const pcFormHeight As Integer = 9000
'Private Const pcFormWidth As Integer = 9720     '640x480
'Private Const pcFormHeight As Integer = 7200
Private mWidth  As Integer
Private mHeight As Integer
Private mTop    As Integer
Private mLeft   As Integer

Private mForm    As Form
Private mPic     As Object
Private mCaption As String

Private Sub Position()
    Dim vPos As String
    vPos = mReg.GetFormPosition(mForm.Caption)
    If vPos <> "" Then
        mForm.Top = Val(Mid(vPos, InStr(vPos, "Top=") + Len("Top=")))
        mForm.Left = Val(Mid(vPos, InStr(vPos, "Left=") + Len("Left=")))
        mForm.Width = Val(Mid(vPos, InStr(vPos, "Width=") + Len("Width=")))
        mForm.Height = Val(Mid(vPos, InStr(vPos, "Height=") + Len("Height=")))
    End If
End Sub

Public Sub MoveSysDate()
    '//lblSysDate ���^�u�ɉB���̂� FrameControl �ō׍H����
    mForm.fraSysDate.BackColor = mForm.lblSysDate.BackColor
    mForm.fraSysDate.Top = mForm.lblSysDate.Top
    mForm.fraSysDate.Left = mForm.lblSysDate.Left
    mForm.fraSysDate.Width = mForm.lblSysDate.Width
    mForm.fraSysDate.Height = mForm.lblSysDate.Height
    mForm.lblSysDate.Top = 0
    mForm.lblSysDate.Left = 0
End Sub

Public Sub Resize()
    On Error Resume Next '//�G���[�����
    '//�T�C�Y�ύX���ɃV�X�e�����t���ړ�����
    mForm.lblSysDate.Left = mForm.Width - (mForm.lblSysDate.Width + 300)
End Sub

Public Sub Init(ByVal vForm As Form, Optional ByVal oDbs As Object = Nothing, Optional oSetMode As Variant = False)
'    On Error Resume Next
    Dim cap As String
    
    Set mForm = vForm
    Call Position
    With mForm
        .Icon = mPic
        mCaption = .Caption
        Call gdDBS.AutoLogOut(mCaption, "Start")
        .Caption = "��" & mReg.CompanyName & "��" & mReg.Title & "-" & mCaption
'''        If oblSetMode = True Then
'''            .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
'''        Else
'''            .Width = pcFormWidth
'''            .Height = pcFormHeight
'''            .Top = pcFormTop
'''            .Left = pcFormLeft
'''        End If
        Dim obj As Object
        For Each obj In .Controls
'            If TypeOf obj Is ORADCLib.ORADC Then
            If TypeName(obj) = "ORADC" Then
                If UCase(TypeName(oDbs)) = UCase("DatabaseClass") Then
                    obj.DatabaseName = oDbs.DatabaseName
                    obj.Connect = oDbs.Connect
'                    obj.Options = dbOption.ORADB_NOWAIT
                    obj.Options = OracleConstantModule.ORADB_NOWAIT
                End If
            End If
        Next obj
        'SysDate �̈ʒu�␳������̂ōŌ�ł���
        .ScaleMode = vbTwips    '//�������Ă����Ȃ��ƈʒu���ςɂȂ�
        With .lblSysDate
            .Caption = Format(Now, "yyyy/mm/dd")
''            .Font = "�l�r ����"
''            .FontSize = 12
            .AutoSize = True
            .Top = 0
            .Left = mForm.Width - (.Width + 300)
        End With
    End With
    On Error GoTo 0
End Sub

Public Sub pInitControl()
    Dim obj As Object
    '// imText,imDate ���N���A����
    For Each obj In mForm.Controls
        Select Case TypeName(obj)
        Case "imText"
            obj.Text = ""
            obj.LengthAsByte = True '//2007/06/06 imText �͑S�ăo�C�g�P�ʂ̓��͂ɂ���
        Case "imNumber"
            obj.Text = ""
        Case "imDate"
            obj.Number = 0
        End Select
    Next obj
End Sub

Public Sub LockedControlAllTextBox()
'//�S�Ă̓��͍��ڂ����b�N����.
    Dim obj As Object
    '//Control.Tag �� "InputKey" ������ݒ肵�Ďg�p�E�s���R���g���[������B
    For Each obj In mForm.Controls
        '�g�p�ۂ�ݒ�
        Select Case TypeName(obj)
        '���j���[�� Visible = False �̃R���g���[���Ɋ܂߂Ȃ�
        Case "imText", _
             "imDate", _
             "imNumber", _
             "Label", _
             "Frame", _
             "OptionButton", _
             "CheckBox", _
             "CommandButton", _
             "ComboBox", "DBList", "DBCombo"
            obj.Enabled = False
        End Select
        If TypeOf obj Is Label Then
            If UCase(obj.Name) = UCase("lblSysDate") Then
                obj.Enabled = True
            End If
        End If
    Next obj
End Sub

Public Sub LockedControl(blMode As Boolean)
    Dim obj As Object
    '//Control.Tag �� "InputKey" ������ݒ肵�Ďg�p�E�s���R���g���[������B
    For Each obj In mForm.Controls
        '�g�p�ۂ�ݒ�
        Select Case TypeName(obj)
        '���j���[�� Visible = False �̃R���g���[���Ɋ܂߂Ȃ�
        Case "imText", _
             "imDate", _
             "imNumber", _
             "Label", _
             "Frame", _
             "OptionButton", _
             "CheckBox", _
             "ComboBox", "DBList", "DBCombo"
'"CommandButton", _�@�R�}���h�{�^���͐��䂩��O���F�{�^���� �`���`�� ����I
            obj.Enabled = ((UCase(obj.Tag) = UCase("InputKey")) = blMode)
        End Select
        '�F��ݒ�
        Select Case TypeName(obj)
        Case "Label", "Frame", "OptionButton", "CheckBox" ', "imDate"
            If obj.BackColor = vbRed Then
                obj.Visible = mReg.Debuged      '//���W�X�g���ݒ�
            ElseIf obj.BackColor <> mForm.BackColor Then
                obj.BackColor = mForm.BackColor
            End If
        End Select
        If TypeOf obj Is Label Then
            If UCase(obj.Name) = UCase("lblSysDate") Then
                obj.Enabled = True
            End If
        End If
    Next obj
End Sub

Public Sub KeyDown(vKeyCode As Integer, Optional oShift As Variant)
#If 0 Then
    If vKeyCode = vbKeyReturn Then
        '//Spread�͈ړ������Ȃ�
        If Not (TypeOf Screen.ActiveControl Is vaSpread) Then
            Call PostMessage(Screen.ActiveForm.hWnd, 256, vbKeyTab, 1)
        End If
    End If
#Else
    If vKeyCode = vbKeyReturn Then
        Select Case TypeName(Screen.ActiveControl)
        Case "TextBox", "imText", "imDate", "imNumber", "OptionButton", _
             "ComboBox", "DBCombo"
            Call PostMessage(Screen.ActiveForm.hWnd, 256, vbKeyTab, 1)
        End Select
    End If
#End If
End Sub

Public Sub SelText(rCtl As Object)
    rCtl.SelStart = 0
    rCtl.SelLength = LenB(rCtl.Text)
End Sub

Public Sub DeSelText(rCtl As Object)
    rCtl.SelStart = LenB(rCtl.Text)
    rCtl.SelLength = 0 'LenB(rCtl.Text)
End Sub

Private Sub Class_Initialize()
    Set mPic = frmAbout.picIcon.Picture
End Sub

Private Sub Class_Terminate()
    Call gdDBS.AutoLogOut(mCaption, "End")
    On Error Resume Next
    Call mReg.SetFormPosition( _
                mCaption, _
                "Top=" & mForm.Top & "," & _
                "Left=" & mForm.Left & "," & _
                "Width=" & mForm.Width & "," & _
                "Height=" & mForm.Height _
            )
    Set mReg = Nothing
    '//ORADC �����ׂăN���[�Y
    Dim obj As Object
    For Each obj In mForm.Controls
'        If TypeOf obj Is ORADCLib.ORADC Then
        If TypeName(obj) = "ORADC" Then
            obj.UpdateControls
            Call obj.Close
        End If
    Next obj
End Sub

Public Function NumToDateFormat(vDate As String) As Variant
    '// "20020101" => "2002/01/01" �ɕϊ�
    '//Variant �ŕԂ��Ȃ��� NULL ���� "00:00:00" ���ԋp����Ă��܂�
    On Error GoTo NumToDateFormatError
    NumToDateFormat = DateSerial(Mid(vDate, 1, 4), Mid(vDate, 5, 2), Mid(vDate, 7, 2))
NumToDateFormatError:
End Function

Public Function DateToNumFormat(vDate As Variant) As Long
    '// "2002/01/01" => "20020101" �ɕϊ�
    On Error GoTo DateToNumFormatError
    '// �N�� 29 => 2029,30 => 1930 �Ɣ��f�����
    If Year(CVDate(vDate)) >= 1930 And Year(CVDate(vDate)) <= 2099 Then
        DateToNumFormat = Format(CVDate(vDate), "yyyymmdd")
    End If
DateToNumFormatError:
End Function

