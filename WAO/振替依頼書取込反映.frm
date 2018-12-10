VERSION 5.00
Begin VB.Form frmFurikaeReqReflectionMode 
   Caption         =   "U‘ÖˆË—Š‘‚Ì”½‰f•û–@‚ğ‘I‘ğ"
   ClientHeight    =   2895
   ClientLeft      =   2835
   ClientTop       =   2040
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2895
   ScaleWidth      =   6750
   Begin VB.Frame grpSelect 
      Caption         =   "y ”½‰f•û–@‚Ì‘I‘ğ z"
      Height          =   1740
      Left            =   375
      TabIndex        =   3
      Top             =   225
      Width           =   5940
      Begin VB.OptionButton optSelect 
         BackColor       =   &H000000FF&
         Caption         =   "‚·‚×‚Ä‚ÌUˆË—Š‘‚Ì“à—e‚ğˆêŠ‡”½‰f‚·‚éB"
         Height          =   315
         Index           =   2
         Left            =   5250
         TabIndex        =   7
         Top             =   1275
         Visible         =   0   'False
         Width           =   3990
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "V‹K“o˜^‚ÌUˆË—Š‘‚Ì‚İ‚ğ y’Ç‰Áz ‚·‚éB"
         Height          =   315
         Index           =   3
         Left            =   450
         TabIndex        =   6
         Top             =   300
         Width           =   210
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "UˆË—Š‘‚ÌŒûÀ‘ŠˆáÒ‚Ì‚İ‚ğ á’uŠ·‚¦â ‚·‚éB"
         Height          =   315
         Index           =   0
         Left            =   450
         TabIndex        =   5
         Top             =   750
         Width           =   210
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "UˆË—Š‘‚ÌŒûÀ‘ŠˆáÒ‚Ì‚İ‚ğ y’Ç‰Áz ‚·‚éB"
         Height          =   315
         Index           =   1
         Left            =   450
         TabIndex        =   4
         Top             =   1200
         Width           =   210
      End
      Begin VB.Label lblIndex1 
         AutoSize        =   -1  'True
         Caption         =   "UˆË—Š‘‚ÌŒûÀ‘ŠˆáÒ‚ğ"
         Height          =   180
         Index           =   0
         Left            =   750
         TabIndex        =   16
         Top             =   1275
         Width           =   2130
      End
      Begin VB.Label lblIndex1 
         AutoSize        =   -1  'True
         Caption         =   " y’Ç‰Áz "
         Height          =   180
         Index           =   1
         Left            =   3225
         TabIndex        =   15
         Top             =   1275
         Width           =   660
      End
      Begin VB.Label lblIndex1 
         AutoSize        =   -1  'True
         Caption         =   "‚·‚éB"
         Height          =   180
         Index           =   2
         Left            =   4200
         TabIndex        =   14
         Top             =   1275
         Width           =   450
      End
      Begin VB.Label lblIndex0 
         AutoSize        =   -1  'True
         Caption         =   "UˆË—Š‘‚ÌŒûÀ‘ŠˆáÒ‚ğ"
         Height          =   180
         Index           =   0
         Left            =   750
         TabIndex        =   13
         Top             =   825
         Width           =   2130
      End
      Begin VB.Label lblIndex0 
         AutoSize        =   -1  'True
         Caption         =   " á’uŠ·‚¦â "
         Height          =   180
         Index           =   1
         Left            =   3100
         TabIndex        =   12
         Top             =   825
         Width           =   1005
      End
      Begin VB.Label lblIndex0 
         AutoSize        =   -1  'True
         Caption         =   "‚·‚éB"
         Height          =   180
         Index           =   2
         Left            =   4200
         TabIndex        =   11
         Top             =   825
         Width           =   450
      End
      Begin VB.Label lblIndex3 
         AutoSize        =   -1  'True
         Caption         =   "‚·‚éB"
         Height          =   180
         Index           =   2
         Left            =   4200
         TabIndex        =   10
         Top             =   375
         Width           =   450
      End
      Begin VB.Label lblIndex3 
         AutoSize        =   -1  'True
         Caption         =   " y’Ç‰Áz "
         Height          =   180
         Index           =   1
         Left            =   3225
         TabIndex        =   9
         Top             =   375
         Width           =   660
      End
      Begin VB.Label lblIndex3 
         AutoSize        =   -1  'True
         Caption         =   "V‹K“o˜^‚ÌUˆË—Š‘‚Ì‚İ‚ğ"
         Height          =   180
         Index           =   0
         Left            =   750
         TabIndex        =   8
         Top             =   375
         Width           =   2310
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "ŠJn(&E)"
      Height          =   495
      Left            =   3375
      TabIndex        =   1
      Top             =   2175
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "’†~(&C)"
      Height          =   510
      Left            =   5025
      TabIndex        =   0
      Top             =   2175
      Width           =   1335
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   195
      Left            =   4500
      TabIndex        =   2
      Top             =   0
      Width           =   1275
   End
End
Attribute VB_Name = "frmFurikaeReqReflectionMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCaption As String
Private mForm As New FormClass
Private mPrevIndex As Integer

Private Sub cmdCancel_Click()
    Call frmFurikaeReqImport.UpdateMode(eModeNon, vbNull)
    Unload Me
End Sub

Private Sub cmdStart_Click()
    If grpSelect.Tag = "" Then
        Call MsgBox("”½‰f•û–@‚ğ‘I‘ğ‚µ‚Ä‰º‚³‚¢B", vbOKOnly + vbInformation, "”½‰f‚ÌŠJn")
        Exit Sub
    End If
    If vbOK <> MsgBox(optSelect(grpSelect.Tag).Caption & vbCrLf & "‚Å" & vbCrLf & "ƒ}ƒXƒ^‚Ì”½‰f‚ğŠJn‚µ‚Ü‚·B" & vbCrLf & vbCrLf & "‚æ‚ë‚µ‚¢‚Å‚·‚©H", vbOKCancel + vbInformation, Me.Caption) Then
        Exit Sub
    End If
    Call frmFurikaeReqImport.UpdateMode(grpSelect.Tag, optSelect(grpSelect.Tag).Caption)
    Unload Me
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    Call OptionClickExtend(lblIndex0)
    Call OptionClickExtend(lblIndex1)
    'Call OptionClickExtend(lblIndex2)
    Call OptionClickExtend(lblIndex3)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mForm = Nothing
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub optSelect_Click(Index As Integer)
    grpSelect.Tag = Index
    
    Dim resetObj As Object
    Select Case mPrevIndex
    Case 0: Set resetObj = lblIndex0
    Case 1: Set resetObj = lblIndex1
    'Case 2
    Case 3: Set resetObj = lblIndex3
    End Select
    mPrevIndex = Index
    Call OptionClickExtend(resetObj, False)
    Set resetObj = Nothing
    
    Dim setObj As Object
    Select Case Index
    Case 0: Set setObj = lblIndex0
    Case 1: Set setObj = lblIndex1
    'Case 2
    Case 3: Set setObj = lblIndex3
    End Select
    Call OptionClickExtend(setObj, True)
    Set setObj = Nothing
End Sub

Private Sub OptionClickExtend(vObj As Object, Optional vSet As Boolean = False)
    If vObj Is Nothing Then
        Exit Sub
    End If
    Dim col_1 As Long, col_2 As Long, col_3 As Long
    If vSet = True Then
        col_1 = vbRed
        If 0 < InStr(vObj(1).Caption, "’Ç‰Á") Then
            col_2 = vbBlue
        Else
            col_2 = vbMagenta
        End If
        col_3 = vbRed
    End If
    vObj(0).FontBold = vSet
    vObj(1).FontBold = vSet
    vObj(2).FontBold = vSet
    vObj(0).ForeColor = col_1
    vObj(1).ForeColor = col_2
    vObj(2).ForeColor = col_3
    vObj(1).Left = vObj(0).Left + vObj(0).Width
    vObj(2).Left = vObj(1).Left + vObj(1).Width
End Sub

Private Sub lblIndex0_Click(Index As Integer)
    optSelect(0).Value = True
End Sub

Private Sub lblIndex1_Click(Index As Integer)
    optSelect(1).Value = True
End Sub

'Private Sub lblIndex2_Click(Index As Integer)
'
'End Sub

Private Sub lblIndex3_Click(Index As Integer)
    optSelect(3).Value = True
End Sub
