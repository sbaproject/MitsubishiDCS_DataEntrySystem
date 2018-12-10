VERSION 5.00
Begin VB.Form frmItakuMasterDataExport 
   Caption         =   "委託者マスタデータ作成"
   ClientHeight    =   3195
   ClientLeft      =   7245
   ClientTop       =   3720
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5640
   Begin VB.TextBox Text4 
      Alignment       =   1  '右揃え
      Height          =   315
      Left            =   2460
      TabIndex        =   3
      Text            =   "2002/07/28"
      Top             =   300
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "作成(&M)"
      Height          =   435
      Left            =   420
      TabIndex        =   2
      Top             =   2220
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "送信(&S)"
      Height          =   435
      Left            =   2160
      TabIndex        =   1
      Top             =   2220
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "終了(&X)"
      Default         =   -1  'True
      Height          =   435
      Left            =   3840
      TabIndex        =   0
      Top             =   2220
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "有効日"
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
      Caption         =   "ﾌｧｲﾙ(&F)"
      Begin VB.Menu mnuEnd 
         Caption         =   "終了(&X)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "ﾍﾙﾌﾟ(&H)"
      Begin VB.Menu mnuVersion 
         Caption         =   "ﾊﾞｰｼﾞｮﾝ情報(&A)"
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
    lblMessage.Caption = "作業手順" & vbCrLf & vbCrLf & "　１：作成処理をします." & vbCrLf & vbCrLf & "作成結果が表示されますので内容に従ってください." & vbCrLf & vbCrLf & "　２：送信処理をします."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsForm = Nothing
    Call gdForm.Show
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub
