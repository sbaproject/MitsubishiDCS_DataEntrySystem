VERSION 5.00
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmKeiyakushaMasterExport 
   Caption         =   "契約者マスタデータ作成"
   ClientHeight    =   3825
   ClientLeft      =   2865
   ClientTop       =   4035
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6180
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
      Caption         =   "Text作成(ホスト向け)(&E)"
      Height          =   555
      Left            =   600
      TabIndex        =   4
      Top             =   3060
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport2 
      Caption         =   "CSV作成(ＤＦ向け) (&S)"
      Height          =   555
      Left            =   2340
      TabIndex        =   5
      Top             =   3060
      Width           =   1455
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "終了(&X)"
      Height          =   555
      Left            =   4020
      TabIndex        =   6
      Top             =   3060
      Width           =   1455
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
   Begin VB.Label lblMessageB 
      Caption         =   "ＤＦ系 メッセージ"
      Height          =   915
      Left            =   780
      TabIndex        =   11
      Top             =   2040
      Width           =   5355
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "【作成手順】"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   900
      TabIndex        =   10
      Top             =   1020
      Width           =   990
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
      TabIndex        =   9
      Top             =   600
      Width           =   915
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label1"
      Height          =   195
      Left            =   4140
      TabIndex        =   8
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label Label8 
      Caption         =   "契約有効日"
      Height          =   255
      Left            =   1380
      TabIndex        =   7
      Top             =   180
      Width           =   915
   End
   Begin VB.Label lblMessageA 
      Caption         =   "ホスト系 メッセージ"
      Height          =   675
      Left            =   780
      TabIndex        =   3
      Top             =   1320
      Width           =   5355
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
Private Const mExeMsgA As String = "ホスト向け Text作成" & vbCrLf & _
                                "　ホストへの固定長テキスト作成をします。" & vbCrLf
Private Const mExeMsgB As String = "ＤＦ向け CSV作成" & vbCrLf & _
                                "　ＤＦへのＣＳＶテキスト作成をします。" & vbCrLf & _
                                "　※注：契約有効日、新規基準日は無視されます。" & vbCrLf
Private mForm As New FormClass
Private mAbort As Boolean

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    '//メッセージ初期化
    lblMessageA.Caption = mExeMsgA
    lblMessageA.FontBold = False
    lblMessageA.ForeColor = vbBlack
    lblMessageB.Caption = mExeMsgB
    lblMessageB.FontBold = False
    lblMessageB.ForeColor = vbBlack
    Const cFileID As String = "ホスト向け."
'    On Error GoTo cmdExport_ClickError
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset, dyn2 As OraDynaset
#Else
    Dim sql As String, dyn As Object, dyn2 As Object
#End If
    
    sql = "SELECT * "
    sql = sql & " FROM taItakushaMaster,"
    sql = sql & "      tbKeiyakushaMaster"
    sql = sql & " WHERE ABITKB = BAITKB"
    '//契約日が有効範囲か？
    sql = sql & "   AND " & txtBAKYED.Number & " BETWEEN BAKYST AND BAKYED"
    '//振替日の有効範囲か？
    sql = sql & "   AND " & txtBAKYED.Number & " BETWEEN BAFKST AND BAFKED"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If dyn.EOF Then
        Call MsgBox("該当するデータはありません.", vbInformation + vbOKOnly, cFileID & mCaption)
        Exit Sub
    End If
    Dim st As New StructureClass, tmp As String
    Dim reg As New RegistryClass
    Dim mFile As New FileClass, FileName As String, TmpFname As String
    Dim fp As Integer, cnt As Long
    
    dlgFile.DialogTitle = "名前を付けて保存(" & cFileID & mCaption & ")"
    dlgFile.FileName = reg.OutputFileName(cFileID & mCaption)
    If IsEmpty(mFile.SaveDialog(dlgFile)) Then
        Exit Sub
    End If
    
    Dim ms As New MouseClass
    Call ms.Start
    
    reg.OutputFileName(cFileID & mCaption) = dlgFile.FileName
    Call st.SelectStructure(st.Keiyakusha)
    
    '//取り敢えずテンポラリに書く
    fp = FreeFile
    TmpFname = mFile.MakeTempFile
    Open TmpFname For Append As #fp
    Do Until dyn.EOF
        DoEvents
        If mAbort Then
            GoTo cmdExport_ClickAbort
        End If
        tmp = ""
        tmp = tmp & st.SetData(dyn.Fields("ABITCD"), 0)     '委託者番号             '//この項目は委託者マスタ
        tmp = tmp & st.SetData(dyn.Fields("BAKYCD"), 1)     '契約者番号(教室)
'//2002/12/10 教室区分(??KSCD)は使用しない
'//        tmp = tmp & st.SetData(dyn.Fields("BAKSCD"), 2)     '教室区分
        tmp = tmp & st.SetData("000", 2)     '教室区分
        '//2002/11/26 空白５文字追加
        tmp = tmp & String(5, " ")
        '//金融機関の区分によって銀行か郵便局の結果を返却する関数を StructureClass を作成
        tmp = tmp & st.SetData(st.BankCode(dyn), 3)         '銀行コード
        tmp = tmp & st.SetData(st.ShitenCode(dyn), 4)       '支店コード
        tmp = tmp & st.SetData(st.Shubetsu(dyn), 5)         '預金種目
        tmp = tmp & st.SetData(st.KouzaNo(dyn), 6)          '口座番号
        '//金融機関の区分によって銀行か郵便局の結果を返却する関数を StructureClass を作成
        tmp = tmp & st.SetData(dyn.Fields("BAKZNM"), 7)     '口座名義人名(カナ)
        tmp = tmp & st.SetData(dyn.Fields("BAZPC1"), 8)     '郵便番号１
        tmp = tmp & st.SetData(dyn.Fields("BAZPC2"), 9)     '郵便番号２
        tmp = tmp & st.SetData(dyn.Fields("BAADJ1"), 10)    '住所１(漢字)
        tmp = tmp & st.SetData(dyn.Fields("BAADJ2"), 11)    '住所２(漢字)
        tmp = tmp & st.SetData(dyn.Fields("BAADJ3"), 12)    '住所３(漢字)
        tmp = tmp & st.SetData(dyn.Fields("BAKJNM"), 13)    '氏名
        tmp = tmp & st.SetData(dyn.Fields("BAKSNO"), 14)    '教室番号
        tmp = tmp & st.SetData(dyn.Fields("BATELE"), 15)    '電話番号１     (教室)
        tmp = tmp & st.SetData(dyn.Fields("BATELJ"), 16)    '電話番号２     (自宅)
        tmp = tmp & st.SetData(dyn.Fields("BAKKRN"), 17)    '電話番号３     (緊急)
        tmp = tmp & st.SetData(dyn.Fields("BAFAXI"), 18)    'ＦＡＸ番号１   (教室)
        tmp = tmp & st.SetData(dyn.Fields("BAFAXJ"), 19)    'ＦＡＸ番号２   (自宅)
        '//保護者の新規人数をカウントする
        sql = "SELECT COUNT(*) AS CNT"
        sql = sql & " FROM tcHogoshaMaster"
        sql = sql & " WHERE CAITKB = '" & CStr(dyn.Fields("BAITKB")) & "'"
        sql = sql & "   AND CAKYCD = '" & CStr(dyn.Fields("BAKYCD")) & "'"
'//2002/12/10 教室区分(??KSCD)は使用しない
'//        sql = sql & "   AND CAKSCD = '" & dyn.Fields("BAKSCD") & "'"
        '//契約開始終了が範囲内か？
        sql = sql & "   AND " & txtBAKYED.Number & " BETWEEN CAKYST AND CAKYED"
        '//契約開始日が有効日以下か？
        sql = sql & "   AND CAADDT >= TO_DATE('" & txtNewData.Text & " 00:00:00','YYYY/MM/DD HH24:MI:SS')"
#If ORA_DEBUG = 1 Then
        Set dyn2 = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
        Set dyn2 = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
        tmp = tmp & st.SetData(dyn2.Fields("CNT"), 20)      '保護者の新規人数
'//2003/02/03 (21)の項目が不足していたので追加
        tmp = tmp & st.SetData(0, 21)                       '調整額：何を入れるの？？？
        Call dyn2.Close
        Print #fp, tmp
        cnt = cnt + 1
        Call dyn.MoveNext
    Loop
    Call dyn.Close
#If NO_TOTAL_REC Then
'//2003/02/03 ここから　合計件数・金額レコード追加
    tmp = ""
    tmp = tmp & st.SetData("9999999999", 0)     '委託者番号             '//この項目は委託者マスタ
    tmp = tmp & st.SetData("9999999999", 1)     '契約者番号(教室)
    tmp = tmp & st.SetData("9999999999", 2)     '教室区分
    tmp = tmp & String(5, " ")                  '//2002/11/26 空白５文字追加
    tmp = tmp & st.SetData("", 3)               '銀行コード
    tmp = tmp & st.SetData("", 4)               '支店コード
    tmp = tmp & st.SetData("", 5)               '預金種目
    tmp = tmp & st.SetData("", 6)               '口座番号
    tmp = tmp & st.SetData("", 7)               '口座名義人名(カナ)
    tmp = tmp & st.SetData("", 8)               '郵便番号１
    tmp = tmp & st.SetData("", 9)               '郵便番号２
    tmp = tmp & st.SetData("合計レコード", 10)  '住所１(漢字)
    tmp = tmp & st.SetData("", 11)              '住所２(漢字)
    tmp = tmp & st.SetData("", 12)              '住所３(漢字)
    tmp = tmp & st.SetData("", 13)              '氏名
    tmp = tmp & st.SetData("", 14)              '教室番号
    tmp = tmp & st.SetData("", 15)              '電話番号１     (教室)
    tmp = tmp & st.SetData("", 16)              '電話番号２     (自宅)
    tmp = tmp & st.SetData("", 17)              '電話番号３     (緊急)
    tmp = tmp & st.SetData("", 18)              'ＦＡＸ番号１   (教室)
    tmp = tmp & st.SetData("", 19)              'ＦＡＸ番号２   (自宅)
    tmp = tmp & st.SetData("", 20)              '保護者の新規人数
    tmp = tmp & st.SetData(cnt, 21)             '＠＠＠ 合計金額 ＠＠＠ 保護者の新規人数
    Print #fp, tmp
'//2003/02/03 ここまで　合計件数・金額レコード追加
#End If
    Close #fp
#If 1 Then
    '//ファイル移動     MOVEFILE_REPLACE_EXISTING=Replace , MOVEFILE_COPY_ALLOWED=Copy & Delete
    Call MoveFileEx(TmpFname, reg.OutputFileName(cFileID & mCaption), MOVEFILE_REPLACE_EXISTING + MOVEFILE_COPY_ALLOWED)
    'Call MoveFileEx(TmpFname, reg.FileName(cFileID & mCaption), MOVEFILE_REPLACE_EXISTING)
#Else
    '//ファイルコピー
    Call FileCopy(TmpFname, reg.FileName(cFileID & mCaption))
#End If
    Set mFile = Nothing
    lblMessageA.Caption = mExeMsgA & cnt & " 件のデータが作成されました。"
    lblMessageA.FontBold = True
    lblMessageA.ForeColor = vbBlue
    Exit Sub
cmdExport_ClickAbort:
    lblMessageA.Caption = mExeMsgA & "中止されました。"
    lblMessageA.FontBold = True
    lblMessageA.ForeColor = vbRed
    Exit Sub
cmdExport_ClickError:
    Call gdDBS.ErrorCheck       '//エラートラップ
    Set mFile = Nothing
End Sub

Private Sub cmdExport2_Click()
    '//メッセージ初期化
    lblMessageA.Caption = mExeMsgA
    lblMessageA.FontBold = False
    lblMessageA.ForeColor = vbBlack
    lblMessageB.Caption = mExeMsgB
    lblMessageB.FontBold = False
    lblMessageB.ForeColor = vbBlack
    Const cFileID As String = "ＤＦ向け."
    'On Error GoTo cmdExport2_ClickError
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset, dyn2 As OraDynaset
#Else
    Dim sql As String, dyn As Object, dyn2 As Object
#End If
    
    sql = "SELECT * "
    sql = sql & " FROM taItakushaMaster,"
    sql = sql & "      tbKeiyakushaMaster"
    sql = sql & " WHERE ABITKB = BAITKB"
    sql = sql & "   AND (BAITKB,BAKYCD,BASQNO) IN("
    sql = sql & "       SELECT BAITKB,BAKYCD,max(BASQNO)"
    sql = sql & "       FROM tbKeiyakushaMaster"
    sql = sql & "       GROUP BY BAITKB,BAKYCD"
    sql = sql & "   )"
    sql = sql & "   AND BAKJNM IS NOT NULL"     '//契約者氏名の NULL があるので排除
    sql = sql & "   AND BAKNNM IS NOT NULL"     '//契約者氏名の NULL があるので排除
'    '//契約日が有効範囲か？
'    sql = sql & "   AND " & txtBAKYED.Number & " BETWEEN BAKYST AND BAKYED"
'    '//振替日の有効範囲か？
'    sql = sql & "   AND " & txtBAKYED.Number & " BETWEEN BAFKST AND BAFKED"
    sql = sql & " ORDER BY BAITKB,BAKYCD"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If dyn.EOF Then
        Call MsgBox("該当するデータはありません.", vbInformation + vbOKOnly, cFileID & mCaption)
        Exit Sub
    End If
    
    Dim reg As New RegistryClass
    Dim mFile As New FileClass, FileName As String, TmpFname As String
    Dim fp As Integer, cnt As Long
    
    dlgFile.DialogTitle = "名前を付けて保存(" & cFileID & mCaption & ")"
    dlgFile.FileName = reg.OutputFileName(cFileID & mCaption)
    If IsEmpty(mFile.SaveDialog(dlgFile)) Then
        Exit Sub
    End If
    
    Dim ms As New MouseClass
    Call ms.Start
    
    reg.OutputFileName(cFileID & mCaption) = dlgFile.FileName
    '//取り敢えずテンポラリに書く
    fp = FreeFile
    TmpFname = mFile.MakeTempFile
    Open TmpFname For Append As #fp
    Do Until dyn.EOF
        DoEvents
        If mAbort Then
            GoTo cmdExport2_ClickAbort
        End If
        '//住所は１、２、３の項目を結合して１２０バイトで返却
        Write #fp, dyn.Fields("BAKYCD"), _
                   dyn.Fields("BAKJNM"), _
                   dyn.Fields("BAKNNM"), _
                   pJoinStrings(dyn.Fields("BAADJ1") & dyn.Fields("BAADJ2") & dyn.Fields("BAADJ3"), 120)
        cnt = cnt + 1
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    Close #fp
#If 1 Then
    '//ファイル移動     MOVEFILE_REPLACE_EXISTING=Replace , MOVEFILE_COPY_ALLOWED=Copy & Delete
    Call MoveFileEx(TmpFname, reg.OutputFileName(cFileID & mCaption), MOVEFILE_REPLACE_EXISTING + MOVEFILE_COPY_ALLOWED)
    'Call MoveFileEx(TmpFname, reg.FileName(cFileID & mCaption), MOVEFILE_REPLACE_EXISTING)
#Else
    '//ファイルコピー
    Call FileCopy(TmpFname, reg.FileName(cFileID & mCaption))
#End If
    Set mFile = Nothing
    lblMessageB.Caption = mExeMsgB & cnt & " 件のデータが作成されました。"
    lblMessageB.FontBold = True
    lblMessageB.ForeColor = vbMagenta

    Exit Sub
cmdExport2_ClickAbort:
    lblMessageB.Caption = mExeMsgA & "中止されました。"
    lblMessageB.FontBold = True
    lblMessageB.ForeColor = vbRed
    Exit Sub
cmdExport2_ClickError:
    Call gdDBS.ErrorCheck       '//エラートラップ
    Set mFile = Nothing
End Sub

'// Variant で受けないと DBNull でエラーとなる
Private Function pJoinStrings(vString As Variant, vBytes As Integer) As String
    If 0 <> Len(vString) Then
        '//全ての半角全角スペースが削除される
        'pJoinStrings = Trim(StrConv(LeftB(StrConv(Replace(Replace(vString, "　", ""), " ", ""), vbFromUnicode), vBytes), vbUnicode))
        '//前後の半角全角スペースが削除される
        pJoinStrings = Trim(StrConv(LeftB(StrConv(vString, vbFromUnicode), vBytes), vbUnicode))
    End If
End Function


'Private Sub cmdSend_Click()
'    Dim reg As New RegistryClass
'    Call Shell(reg.TransferCommand(mCaption), vbNormalFocus)
'End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    cmdExport.Caption = "ホスト向け" & vbCrLf & "Text作成(&H)"
    cmdExport2.Caption = "ＤＦ向け" & vbCrLf & "CSV作成(&D)"
    lblMessageA.Caption = mExeMsgA
    lblMessageB.Caption = mExeMsgB
    txtBAKYED.Number = gdDBS.sysDate("YYYYMMDD")
    txtNewData.Number = gdDBS.sysDate("YYYYMMDD")
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mAbort = True
    Set frmKeiyakushaMasterExport = Nothing
    Set mForm = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
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

