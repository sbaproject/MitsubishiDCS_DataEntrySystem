VERSION 5.00
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Begin VB.Form frmMakeNewData 
   Caption         =   "�����f�[�^�ǉ�"
   ClientHeight    =   3300
   ClientLeft      =   3750
   ClientTop       =   1800
   ClientWidth     =   6345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6345
   Begin VB.CommandButton cmdReturn 
      Caption         =   "�㏑��(&U)"
      Height          =   435
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      ToolTipText     =   "������ǉ����Ȃ��ł��̂܂܍X�V����ꍇ"
      Top             =   2580
      Width           =   1395
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "�ǉ�(&A)"
      Height          =   435
      Index           =   1
      Left            =   480
      TabIndex        =   1
      ToolTipText     =   "������ǉ�����ꍇ"
      Top             =   2580
      Width           =   1395
   End
   Begin VB.CommandButton cmdReturn 
      Cancel          =   -1  'True
      Caption         =   "���~(&C)"
      Height          =   435
      Index           =   0
      Left            =   4740
      TabIndex        =   0
      ToolTipText     =   "���̍�Ƃ𒆎~���čēx���Ƃ̉�ʂ�ҏW����ꍇ"
      Top             =   2580
      Width           =   1335
   End
   Begin imDate6Ctl.imDate txtKeiyakuEnd 
      DataField       =   "BAKYED"
      Height          =   315
      Left            =   3780
      TabIndex        =   4
      Top             =   1620
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   556
      Calendar        =   "�ǉ��f�[�^�쐬.frx":0000
      Caption         =   "�ǉ��f�[�^�쐬.frx":0186
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ǉ��f�[�^�쐬.frx":01F4
      Keys            =   "�ǉ��f�[�^�쐬.frx":0212
      MouseIcon       =   "�ǉ��f�[�^�쐬.frx":0270
      Spin            =   "�ǉ��f�[�^�쐬.frx":028C
      AlignHorizontal =   2
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   1
      DisplayFormat   =   "yyyy/mm/dd"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy/mm/dd"
      HighlightText   =   2
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   " "
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "    /  /  "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   -2
      CenturyMode     =   0
   End
   Begin imDate6Ctl.imDate txtFurikaeEnd 
      DataField       =   "BAFKED"
      Height          =   315
      Left            =   3780
      TabIndex        =   5
      Top             =   2040
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   556
      Calendar        =   "�ǉ��f�[�^�쐬.frx":02B4
      Caption         =   "�ǉ��f�[�^�쐬.frx":043A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ǉ��f�[�^�쐬.frx":04A8
      Keys            =   "�ǉ��f�[�^�쐬.frx":04C6
      MouseIcon       =   "�ǉ��f�[�^�쐬.frx":0524
      Spin            =   "�ǉ��f�[�^�쐬.frx":0540
      AlignHorizontal =   2
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   1
      DisplayFormat   =   "yyyy/mm/dd"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy/mm/dd"
      HighlightText   =   2
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   " "
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "    /  /  "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   -2
      CenturyMode     =   0
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
      Caption         =   "�ǉ������f�[�^��"
      Height          =   255
      Left            =   1140
      TabIndex        =   9
      Top             =   1680
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�ǉ������f�[�^��"
      Height          =   255
      Left            =   1140
      TabIndex        =   8
      Top             =   2100
      Width           =   1515
   End
   Begin VB.Label lblFurikomi 
      Alignment       =   1  '�E����
      Caption         =   "�U���J�n��"
      Height          =   255
      Left            =   2700
      TabIndex        =   7
      Top             =   2100
      Width           =   975
   End
   Begin VB.Label Label19 
      Alignment       =   1  '�E����
      Caption         =   "�L���J�n��"
      Height          =   255
      Left            =   2700
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblMessage 
      Caption         =   "lblMessage"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   6015
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
Attribute VB_Name = "frmMakeNewData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm As New FormClass

'//�߂�t�H�[���ŎQ�Ƃ���ϐ�
Public mPushButton As Integer
Public Enum ePushButton
    Cancel = 0
    Add = 1
    Update = 2
End Enum
Public mKeiyakuEnd As Long
Public mFurikaeEnd As Long

Private Sub cmdReturn_Click(Index As Integer)
'    lblPushButton.Caption = Index   '//�I�u�W�F�N�g���쐬���Ă�����Ƃ��ɔj�������
    mPushButton = Index             '//����������ϐ��͑���-Form �ɕύX������ԂŌ�����
 '''//2002/10/18 ���̂܂܂̓��t�Ƃ���
'''   '//�N���݂̂̓��͂Ȃ̂� 2/31 �Ƃ������݂��邽��
'''    mKeiyakuEnd = Format(DateSerial(txtKeiyakuEnd.Year, txtKeiyakuEnd.Month, 1), "yyyymmdd")
'''    mFurikaeEnd = Format(DateSerial(txtFurikaeEnd.Year, txtFurikaeEnd.Month, 1), "yyyymmdd")
    mKeiyakuEnd = txtKeiyakuEnd.Number
    mFurikaeEnd = txtFurikaeEnd.Number
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
'    Call Me.Move((Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2)
    Me.Height = 4000    '//�X�^�[�g���j���[�ɍ��E����ăT�C�Y�����������Ȃ�̂ŋ����I�ɐݒ肷��.
    Me.Icon = frmAbout.Icon
'    lblMessage.Caption = "��Ǝ菇" & vbCrLf & vbCrLf & "�@�P�F��M���������܂�." & vbCrLf & "�@�Q�F�捞���������܂�." & vbCrLf & vbCrLf & "�捞���ʂ��\������܂��̂œ��e�ɏ]���Ă�������."
    txtKeiyakuEnd.Number = gdDBS.sysDate("YYYYMMDD")
    txtFurikaeEnd.Number = gdDBS.sysDate("YYYYMMDD")
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMakeNewData = Nothing
    Set mForm = Nothing
'    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub mnuEnd_Click()
    Call cmdReturn_Click(ePushButton.Cancel)
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

Private Sub txtFurikaeEnd_DropOpen(NoDefault As Boolean)
    txtFurikaeEnd.Calendar.Holidays = gdDBS.Holiday(txtFurikaeEnd.Year)
End Sub

Private Sub txtKeiyakuEnd_DropOpen(NoDefault As Boolean)
    txtKeiyakuEnd.Calendar.Holidays = gdDBS.Holiday(txtKeiyakuEnd.Year)
End Sub

