VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Begin VB.Form frmClassNoChange 
   Caption         =   "�����ԍ��̕ύX"
   ClientHeight    =   2520
   ClientLeft      =   3525
   ClientTop       =   2235
   ClientWidth     =   4575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4575
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�X�V(&U)"
      Height          =   435
      Left            =   480
      TabIndex        =   1
      ToolTipText     =   "������ǉ�����ꍇ"
      Top             =   1860
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "���~(&C)"
      Height          =   435
      Left            =   2700
      TabIndex        =   2
      ToolTipText     =   "���̍�Ƃ𒆎~���čēx���Ƃ̉�ʂ�ҏW����ꍇ"
      Top             =   1860
      Width           =   1335
   End
   Begin imText6Ctl.imText txtCAKSCD 
      Height          =   285
      Left            =   2820
      TabIndex        =   0
      Tag             =   "InputKey"
      Top             =   1380
      Width           =   375
      _Version        =   65537
      _ExtentX        =   661
      _ExtentY        =   503
      Caption         =   "�����ԍ��ύX.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�����ԍ��ύX.frx":006C
      Key             =   "�����ԍ��ύX.frx":008A
      MouseIcon       =   "�����ԍ��ύX.frx":00CE
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   0
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   3
      LengthAsByte    =   -1
      Text            =   "123"
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   3
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.Label Label3 
      Caption         =   "���ύX���e�͑����X�V����܂��B"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblCAKSCD 
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      Caption         =   "001"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2820
      TabIndex        =   9
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblCAHGCD 
      Alignment       =   1  '�E����
      BackColor       =   &H000000FF&
      Caption         =   "�ی��"
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCAKYCD 
      Alignment       =   1  '�E����
      BackColor       =   &H000000FF&
      Caption         =   "�_���"
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCAITKB 
      Alignment       =   1  '�E����
      BackColor       =   &H000000FF&
      Caption         =   "�ϑ���"
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
      Caption         =   "�ύX�O�̋����ԍ�"
      Height          =   255
      Left            =   1140
      TabIndex        =   5
      Top             =   960
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�ύX��̋����ԍ�"
      Height          =   255
      Left            =   1140
      TabIndex        =   4
      Top             =   1380
      Width           =   1515
   End
   Begin VB.Label lblMessage 
      Caption         =   "�Ԉ���ē��͂��ꂽ�����ԍ��̕ύX�����܂��B"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1020
      TabIndex        =   3
      Top             =   120
      Width           =   2460
   End
   Begin VB.Menu mnuFile 
      Caption         =   "̧��(&F)"
      Begin VB.Menu mnuEnd 
         Caption         =   "�I��(&X)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuVersion 
         Caption         =   "�ް�ޮݏ��(&A)"
      End
   End
End
Attribute VB_Name = "frmClassNoChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm As New FormClass
Public mNewCode As String
Private mIsActivated As Boolean

'//2007/06/07 �X�V�E���~�{�^�������S�P�ƂɃR���g���[��
Private Sub pButtonControl(ByVal vMode As Boolean, Optional vExec As Boolean = False)
    If True = mIsActivated Or True = vExec Then
        cmdUpdate.Visible = vMode
        'cmdCancel.Visible = vMode
        cmdUpdate.Enabled = vMode
        'cmdCancel.Enabled = vMode
        mIsActivated = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    If True = pUpdateCheck Then
        Unload Me
    End If
End Sub

Private Function pUpdateCheck() As Boolean
    If "" = Trim(txtCAKSCD.Text) Then
        Call MsgBox("�����ԍ��������͂ł�.", vbCritical + vbOKOnly, Me.Caption)
        Exit Function
    ElseIf Trim(lblCAKSCD.Caption) = Trim(txtCAKSCD.Text) Then
        Call MsgBox("���������ԍ��ɂ͕ύX�ł��܂���.", vbCritical + vbOKOnly, Me.Caption)
        Exit Function
    End If
    '//�e�}�X�^�E�g�����̑��݃`�F�b�N������
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    '//�Q�񍻎��v���o���̂� New() ���Ȃ�
    Dim ms As MouseClass
    Set ms = New MouseClass
    Call ms.Start
    '//�ی�҃}�X�^
    sql = "SELECT CAHGCD FROM tcHogoshaMaster"
    sql = sql & " WHERE CAITKB = '" & lblCAITKB.Caption & "'"
    sql = sql & "   AND CAKYCD = '" & lblCAKYCD.Caption & "'"
    sql = sql & "   AND CAKSCD = '" & Trim(txtCAKSCD.Text) & "'" '//���̃R�[�h�̑��݃`�F�b�N
    sql = sql & "   AND CAHGCD = '" & lblCAHGCD.Caption & "'"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If Not dyn.EOF Then
        Call pUpdateErrorMsg("�ی�҃}�X�^", dyn.Fields("CAHGCD").Value)
        Exit Function
    End If
    Call dyn.Close
    '//�����U�փf�[�^�F�ݐς͍X�V���Ȃ�(���̂܂܂ɂ��Ă���)
    sql = "SELECT FAHGCD FROM tfFurikaeYoteiData"
    sql = sql & " WHERE FAITKB = '" & lblCAITKB.Caption & "'"
    sql = sql & "   AND FAKYCD = '" & lblCAKYCD.Caption & "'"
    sql = sql & "   AND FAKSCD = '" & Trim(txtCAKSCD.Text) & "'" '//���̃R�[�h�̑��݃`�F�b�N
    sql = sql & "   AND FAHGCD = '" & lblCAHGCD.Caption & "'"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If Not dyn.EOF Then
        Call pUpdateErrorMsg("�����U�փf�[�^", dyn.Fields("FAHGCD").Value)
        Exit Function
    End If
    Call dyn.Close
    Set ms = Nothing
    If vbOK <> MsgBox("�����ԍ��� (" & lblCAKSCD.Caption & ") ���� (" & txtCAKSCD.Text & ") �ɕύX���܂����H", vbInformation + vbOKCancel, Me.Caption) Then
        Exit Function
    End If
    '//�`�F�b�N�n�j�F���ɑ��݂��Ȃ��̂ōX�V�\
    Set ms = New MouseClass
    Call ms.Start
    On Error GoTo pUpdateCheckError
    '�r����������̂ŕی�҃}�X�^�E���b�N����
    Call frmHogoshaMaster.dbcHogoshaMaster.UpdateControls
    Call gdDBS.Database.BeginTrans
    '//�ی�҃}�X�^
    sql = "UPDATE tcHogoshaMaster SET "
    sql = sql & " CAKSCD = '" & Trim(txtCAKSCD.Text) & "',"
    sql = sql & " CAUSID = '" & gdDBS.LoginUserName & "',"
    sql = sql & " CAUPDT = SYSDATE"
    sql = sql & " WHERE CAITKB = '" & lblCAITKB.Caption & "'"
    sql = sql & "   AND CAKYCD = '" & lblCAKYCD.Caption & "'"
    sql = sql & "   AND CAKSCD = '" & lblCAKSCD.Caption & "'"  '//KEY:���̃R�[�h��ύX
    sql = sql & "   AND CAHGCD = '" & lblCAHGCD.Caption & "'"
    Call gdDBS.Database.ExecuteSQL(sql)
    '//�����U�փf�[�^�F�ݐς͍X�V���Ȃ�(���̂܂܂ɂ��Ă���)
    sql = "UPDATE tfFurikaeYoteiData SET "
    sql = sql & " FAKSCD = '" & Trim(txtCAKSCD.Text) & "',"
    sql = sql & " FAUSID = '" & gdDBS.LoginUserName & "',"
    sql = sql & " FAUPDT = SYSDATE"
    sql = sql & " WHERE FAITKB = '" & lblCAITKB.Caption & "'"
    sql = sql & "   AND FAKYCD = '" & lblCAKYCD.Caption & "'"
    sql = sql & "   AND FAKSCD = '" & lblCAKSCD.Caption & "'"  '//KEY:���̃R�[�h��ύX
    sql = sql & "   AND FAHGCD = '" & lblCAHGCD.Caption & "'"
    Call gdDBS.Database.ExecuteSQL(sql)
    
    Call gdDBS.Database.CommitTrans
    mNewCode = txtCAKSCD.Text
    pUpdateCheck = True
    Exit Function
pUpdateCheckError:
    Call gdDBS.Database.Rollback
    Call gdDBS.ErrorCheck       '//�G���[�g���b�v
'// gdDBS.ErrorCheck() �̏�Ɉړ�
'//    Call gdDBS.Database.Rollback
    Call gdDBS.AutoLogOut(Me.Caption, " �G���[�������������ߏ����͒��~����܂����B")
End Function

Private Sub pUpdateErrorMsg(vMst As String, vCode As String)
    Call MsgBox(vMst & "�� (" & vCode & ")�̐l��" & vbCrLf & "���݂��邽�ߕύX�͏o���܂���.", vbCritical + vbOKOnly, Me.Caption)
End Sub

Private Sub Form_Activate()
    If False = mIsActivated Then
        Call pButtonControl(False, True)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    Me.Height = 3200      '//�X�^�[�g���j���[�ɍ��E����ăT�C�Y�����������Ȃ�̂ŋ����I�ɐݒ肷��.
    Me.Icon = frmAbout.Icon
    mNewCode = ""
End Sub

#If 0 Then
Private Sub Form_Resize()
    Call mForm.Resize
End Sub
#End If

Private Sub Form_Unload(Cancel As Integer)
    Set frmClassNoChange = Nothing
    Set mForm = Nothing
'    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub mnuEnd_Click()
    Call cmdCancel_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

Private Sub txtCAKSCD_Change()
    Call pButtonControl(True)
End Sub

Private Sub txtCAKSCD_LostFocus()
'//2006/04/26 �O�[�����ߍ���
    If "" <> Trim(txtCAKSCD.Text) Then
        txtCAKSCD.Text = Format(Val(txtCAKSCD.Text), "000")
    End If
End Sub
