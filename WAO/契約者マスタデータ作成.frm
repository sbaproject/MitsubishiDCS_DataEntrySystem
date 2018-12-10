VERSION 5.00
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmKeiyakushaMasterExport 
   Caption         =   "オーナーマスタデータ作成"
   ClientHeight    =   3450
   ClientLeft      =   2865
   ClientTop       =   4035
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5880
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   540
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin imDate6Ctl.imDate txtBAKYED 
      Height          =   285
      Left            =   2340
      TabIndex        =   0
      Top             =   120
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   503
      Calendar        =   "契約者マスタデータ作成.frx":0000
      Caption         =   "契約者マスタデータ作成.frx":0186
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "契約者マスタデータ作成.frx":01F4
      Keys            =   "契約者マスタデータ作成.frx":0212
      MouseIcon       =   "契約者マスタデータ作成.frx":0270
      Spin            =   "契約者マスタデータ作成.frx":028C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
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
      MaxDate         =   73050
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   " "
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "    /  /  "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   -2
      CenturyMode     =   0
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "作成(&E)"
      Height          =   435
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "終了(&X)"
      Default         =   -1  'True
      Height          =   435
      Left            =   4020
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin imDate6Ctl.imDate txtNewData 
      Height          =   285
      Left            =   2340
      TabIndex        =   1
      Top             =   540
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   503
      Calendar        =   "契約者マスタデータ作成.frx":02B4
      Caption         =   "契約者マスタデータ作成.frx":043A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "契約者マスタデータ作成.frx":04A8
      Keys            =   "契約者マスタデータ作成.frx":04C6
      MouseIcon       =   "契約者マスタデータ作成.frx":0524
      Spin            =   "契約者マスタデータ作成.frx":0540
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
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
      MaxDate         =   73050
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   " "
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "    /  /  "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   -2
      CenturyMode     =   0
   End
   Begin VB.Label Label2 
      Caption         =   "以降を新規扱いとする。"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "新規基準日"
      Height          =   255
      Left            =   1380
      TabIndex        =   8
      Top             =   600
      Width           =   915
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label1"
      Height          =   195
      Left            =   4140
      TabIndex        =   7
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label Label8 
      Caption         =   "契約有効日"
      Height          =   255
      Left            =   1380
      TabIndex        =   6
      Top             =   180
      Width           =   915
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      Height          =   1755
      Left            =   360
      TabIndex        =   3
      Top             =   900
      Width           =   5175
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
Attribute VB_Name = "frmKeiyakushaMasterExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCaption As String
Private Const mExeMsg As String = "作成処理をします." & vbCrLf & vbCrLf & "作成結果が表示されますので内容に従ってください." & vbCrLf & vbCrLf
Private mForm As New FormClass
Private mAbort As Boolean

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
'    On Error GoTo cmdExport_ClickError
'    Dim sql As String, dyn As OraDynaset, dyn2 As OraDynaset
    Dim sql As String, dyn As Object, dyn2 As Object
    
    sql = "SELECT * "
    sql = sql & " FROM taItakushaMaster,"
    sql = sql & "      tbKeiyakushaMaster"
    sql = sql & " WHERE ABITKB = BAITKB"
    '//契約日が有効範囲か？
    sql = sql & "   AND " & txtBAKYED.Number & " BETWEEN BAKYST AND BAKYED"
    '//振替日の有効範囲か？
    sql = sql & "   AND " & txtBAKYED.Number & " BETWEEN BAFKST AND BAFKED"
    sql = sql & " order by LTRIM(nvl(BAKYNY,'XXX')),BAKYCD"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If dyn.EOF Then
        Call MsgBox("該当するデータはありません.", vbInformation + vbOKOnly, mCaption)
        Exit Sub
    End If
    Dim st As New StructureClass, tmp As String
    Dim reg As New RegistryClass
    Dim mFile As New FileClass, FileName As String, TmpFname As String
    Dim fp As Integer, cnt As Long
    
    dlgFile.DialogTitle = "名前を付けて保存(" & mCaption & ")"
    dlgFile.FileName = reg.OutputFileName(mCaption)
    If IsEmpty(mFile.SaveDialog(dlgFile)) Then
        Exit Sub
    End If
    
    Dim ms As New MouseClass
    Call ms.Start
    
    reg.OutputFileName(mCaption) = dlgFile.FileName
    Call st.SelectStructure(st.Keiyakusha)
    
    '//取り敢えずテンポラリに書く
    fp = FreeFile
    TmpFname = mFile.MakeTempFile
    Open TmpFname For Append As #fp
    Do Until dyn.EOF
        DoEvents
        If mAbort Then
            GoTo cmdExport_ClickError
        End If
        tmp = ""
        tmp = tmp & st.SetData(dyn.Fields("ABITCD"), 0)     '委託者番号             '//この項目は委託者マスタ
        tmp = tmp & st.SetData(dyn.Fields("BAKYCD"), 1)     '契約者番号
        tmp = tmp & st.SetData(String(8, "0"), 2)           'FILLER
        tmp = tmp & st.SetData(dyn.Fields("BAKJNM"), 3)     '氏名
        tmp = tmp & st.SetData(dyn.Fields("BAZPC1"), 4)     '郵便番号１
        tmp = tmp & st.SetData(dyn.Fields("BAZPC2"), 5)     '郵便番号２
        tmp = tmp & st.SetData(dyn.Fields("BAADJ1"), 6)    '住所１(漢字)
        tmp = tmp & st.SetData(dyn.Fields("BAADJ2"), 7)    '住所２(漢字)
        tmp = tmp & st.SetData(dyn.Fields("BATELE"), 8)    '電話番号１
       'tmp = tmp & st.SetData(dyn.Fields("BATELJ"), 9)    '電話番号２
        tmp = tmp & st.SetData(dyn.Fields("BAKKRN"), 9)    '2016/11/17 ホストでは緊急連絡先なので正しく渡す
        tmp = tmp & st.SetData(dyn.Fields("BAkome"), 10)    '校名
        tmp = tmp & st.SetData(st.BankCode(dyn), 11)        '銀行コード
        tmp = tmp & st.SetData(st.ShitenCode(dyn), 12)      '支店コード
        tmp = tmp & st.SetData(st.Shubetsu(dyn), 13)        '預金種目
        tmp = tmp & st.SetData(st.KouzaNo(dyn), 14)         '口座番号
        tmp = tmp & st.SetData(dyn.Fields("BAKZNM"), 15)    '口座名義人名(カナ)
        tmp = tmp & st.SetData(dyn.Fields("BAHJNO"), 16)    '法人番号
        tmp = tmp & st.SetData(dyn.Fields("BAKYNY"), 17)    '名寄せ先契約者番号
        tmp = tmp & st.SetData("", 18)                      'FILLER
        Print #fp, tmp
        cnt = cnt + 1
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    Close #fp
#If 1 Then
    '//ファイル移動     MOVEFILE_REPLACE_EXISTING=Replace , MOVEFILE_COPY_ALLOWED=Copy & Delete
    Call MoveFileEx(TmpFname, reg.OutputFileName(mCaption), MOVEFILE_REPLACE_EXISTING + MOVEFILE_COPY_ALLOWED)
    'Call MoveFileEx(TmpFname, reg.FileName(mCaption), MOVEFILE_REPLACE_EXISTING)
#Else
    '//ファイルコピー
    Call FileCopy(TmpFname, reg.FileName(mCaption))
#End If
    Set mFile = Nothing
    lblMessage.Caption = mExeMsg & cnt & " 件のデータが作成されました。"
    Exit Sub
cmdExport_ClickError:
    Call gdDBS.ErrorCheck       '//エラートラップ
    Set mFile = Nothing
End Sub

Private Sub cmdSend_Click()
    Dim reg As New RegistryClass
    Call Shell(reg.TransferCommand(mCaption), vbNormalFocus)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    lblMessage.Caption = mExeMsg
    txtBAKYED.Number = gdDBS.sysDate("YYYYMMDD")
    txtNewData.Number = gdDBS.sysDate("YYYYMMDD")
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(cancel As Integer)
    mAbort = True
    Set frmKeiyakushaMasterExport = Nothing
    Set mForm = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        cancel = True
    End If
End Sub

Private Sub mnuEnd_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

Private Sub txtBAKYED_DropOpen(NoDefault As Boolean)
    txtBAKYED.Calendar.Holidays = gdDBS.Holiday(txtBAKYED.Year)
End Sub

Private Sub txtNewData_DropOpen(NoDefault As Boolean)
    txtNewData.Calendar.Holidays = gdDBS.Holiday(txtNewData.Year)
End Sub

