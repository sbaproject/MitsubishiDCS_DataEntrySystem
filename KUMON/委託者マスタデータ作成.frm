VERSION 5.00
Begin VB.Form frmItakuMasterDataExport 
   Caption         =   "�ϑ��҃}�X�^�f�[�^�쐬"
   ClientHeight    =   3195
   ClientLeft      =   7245
   ClientTop       =   3720
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5640
   Begin VB.TextBox Text4 
      Alignment       =   1  '�E����
      Height          =   315
      Left            =   2460
      TabIndex        =   3
      Text            =   "2002/07/28"
      Top             =   300
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�쐬(&M)"
      Height          =   435
      Left            =   420
      TabIndex        =   2
      Top             =   2220
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���M(&S)"
      Height          =   435
      Left            =   2160
      TabIndex        =   1
      Top             =   2220
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "�I��(&X)"
      Default         =   -1  'True
      Height          =   435
      Left            =   3840
      TabIndex        =   0
      Top             =   2220
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "�L����"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   360
      Width           =   555
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      Height          =   1335
      Left            =   420
      TabIndex        =   5
      Top             =   720
      Width           =   4815
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
Attribute VB_Name = "frmItakuMasterDataExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsForm As New FormClass

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Me.Move((Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2)
    Me.Icon = frmAbout.Icon
    lblMessage.Caption = "��Ǝ菇" & vbCrLf & vbCrLf & "�@�P�F�쐬���������܂�." & vbCrLf & vbCrLf & "�쐬���ʂ��\������܂��̂œ��e�ɏ]���Ă�������." & vbCrLf & vbCrLf & "�@�Q�F���M���������܂�."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsForm = Nothing
    Call gdForm.Show
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub
