VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmBankDataImport 
   Caption         =   "金融機関データ取込"
   ClientHeight    =   3345
   ClientLeft      =   2805
   ClientTop       =   1830
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5640
   Begin VB.CommandButton cmdImport 
      Caption         =   "取込(&I)"
      Height          =   435
      Left            =   2100
      TabIndex        =   2
      Top             =   2580
      Width           =   1395
   End
   Begin VB.CommandButton cmdRecv 
      Caption         =   "受信(&R)"
      Height          =   435
      Left            =   540
      TabIndex        =   1
      Top             =   2580
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "終了(&X)"
      Default         =   -1  'True
      Height          =   435
      Left            =   3960
      TabIndex        =   0
      Top             =   2580
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   480
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar pgrProgressBar 
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2220
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   255
      Left            =   4020
      TabIndex        =   4
      Top             =   0
      Width           =   1395
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      Height          =   1635
      Left            =   480
      TabIndex        =   3
      Top             =   420
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
Attribute VB_Name = "frmBankDataImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pcBankFileLength As Integer = 69

Private mCaption As String
Private Const mExeMsg As String = "作業手順" & vbCrLf & vbCrLf & "　１：受信処理をします." & vbCrLf & "　２：取込処理をします." & vbCrLf & vbCrLf & "取込結果が表示されますので内容に従ってください." & vbCrLf & vbCrLf
Private mForm As New FormClass
Private mAbort As Boolean

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdImport_Click()
    Dim ms As New MouseClass
    Call ms.Start
    On Error GoTo cmdImport_ClickError
    
    Call gdDBS.Database.BeginTrans
    Dim result   As Long
    Dim BankFile As String, reg As New RegistryClass
    
'    BankFile = reg.InputFileName(mCaption)
    BankFile = "D:\公文教育研究会\全銀マスタ.txt"
    
    Call gdDBS.Database.Parameters.Add("Result", result, paramMode.ORAPARM_OUTPUT, serverType.ORATYPE_NUMBER)
    Call gdDBS.Database.Parameters.Add("BankFile", BankFile, paramMode.ORAPARM_INPUT, serverType.ORATYPE_CHAR)
    Call gdDBS.Database.ExecuteSQL("BEGIN :Result := PKG_BANK_IMPORT.MAIN(:BankFile); END;")
    '//返り値を取得
    result = gdDBS.Database.Parameters("Result").Value
    Call gdDBS.Database.Parameters.Remove("Result")
    Call gdDBS.Database.Parameters.Remove("BankFile")
'    If result < 0 Then
'        Call MsgBox("取り込み中にエラーが発生しました.(Error Code = " & result & ")" & vbCrLf & ErrMsg, vbCritical + vbOKOnly, mCaption)
'        Exit Sub
'    End If
    lblMessage.Caption = mExeMsg & result & " 件のデータが取り込まれました。"
    Call gdDBS.Database.CommitTrans
    Exit Sub
cmdImport_ClickError:
    Call gdDBS.Database.Parameters.Remove("Result")
    Call gdDBS.Database.Parameters.Remove("BankFile")
    Call gdDBS.Database.Rollback
    Call gdDBS.ErrorCheck       '//エラートラップ
End Sub

Private Sub cmdRecv_Click()
    Dim reg As New RegistryClass
    Call Shell(reg.TransferCommand(mCaption), vbNormalFocus)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    lblMessage.Caption = mExeMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mAbort = True
    Set mForm = Nothing
    Set frmBankDataImport = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub


Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub
