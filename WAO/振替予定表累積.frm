VERSION 5.00
Begin VB.Form frmFurikaeDataRuiseki 
   Caption         =   "振替予定表 兼 解約通知書(累積)"
   ClientHeight    =   3255
   ClientLeft      =   2295
   ClientTop       =   2130
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6045
   Begin VB.ListBox lstFurikaeBi 
      Height          =   690
      ItemData        =   "振替予定表累積.frx":0000
      Left            =   2340
      List            =   "振替予定表累積.frx":0016
      Style           =   1  'ﾁｪｯｸﾎﾞｯｸｽ
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "実行(&E)"
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
      Left            =   4140
      TabIndex        =   0
      Top             =   2580
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "口座振替日"
      Height          =   195
      Left            =   1260
      TabIndex        =   4
      Top             =   1740
      Width           =   915
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   0
      Width           =   1395
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      Height          =   1155
      Left            =   360
      TabIndex        =   2
      Top             =   420
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
Attribute VB_Name = "frmFurikaeDataRuiseki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCaption As String
Private mForm As New FormClass

Private Const mExeMsg As String = "振替金額予定表 兼 解約通知書データを累積します。" & vbCrLf & vbCrLf

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdExec_Click()
    Dim sql As String, cnt As Long
    Dim WhereSQL As String, ix As Long, msg As String
    Dim ms As New MouseClass
    Call ms.Start
    
'//2004/05/17 累積件数のログ追加のために日付を退避
    Dim RuisekiDate As String
    '//リストでチェックされたデータを IN 句に...。
    WhereSQL = " WHERE FASQNO IN("
    RuisekiDate = "("
    For ix = 0 To lstFurikaeBi.ListCount - 1
        If lstFurikaeBi.Selected(ix) = True Then
            cnt = cnt + 1
            WhereSQL = WhereSQL & Format(lstFurikaeBi.List(ix), "yyyymmdd") & ","
'//2004/05/17 累積件数のログ追加のために日付を退避
            RuisekiDate = RuisekiDate & lstFurikaeBi.List(ix) & ","
        End If
    Next ix
    WhereSQL = Left(WhereSQL, Len(WhereSQL) - 1) & ")"
'//2004/05/17 累積件数のログ追加のために日付を退避
    RuisekiDate = Left(RuisekiDate, Len(RuisekiDate) - 1) & ")"
    If cnt = 0 Then
        msg = "累積すべきデータはありませんでした。"
        lblMessage.Caption = mExeMsg & msg
        Call MsgBox(msg, vbInformation, mCaption)
        Exit Sub
    End If
    
    On Error GoTo cmdExec_ClickError

'//2003/02/03 更新状態フラグをチェックして警告 追加:0=DB作成,1=予定作成,2=予定取込,3=請求作成
    Dim dyn As Object
    sql = "SELECT FASQNO FROM tfFurikaeYoteiData"
    sql = sql & WhereSQL
    sql = sql & " AND FAUPFG < '" & eKouFuriKubun.SeikyuText & "'"
    sql = sql & " AND NVL(FAKYFG,0) = 0"
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
    If dyn.RecordCount > 0 Then
        msg = "請求データの作成されていないデータが存在します." & vbCrLf & "累積処理を続行しますか？"
        lblMessage.Caption = mExeMsg & msg
        If vbOK <> MsgBox(msg, vbInformation + vbOKCancel + vbDefaultButton2, mCaption) Then
            Exit Sub
        End If
    End If
    
    Call gdDBS.Database.BeginTrans
    
'//2004/06/03 新規扱いした日付を保護者マスタ(CANWDT)に設定：金額「０」は新規としない
    sql = "UPDATE tcHogoshaMaster SET "
    sql = sql & " CANWDT = SYSDATE "
    sql = sql & " WHERE (CAITKB,CAKYCD,CAHGCD) IN("
    sql = sql & "       SELECT FAITKB,FAKYCD,FAHGCD "
    sql = sql & "       FROM tfFurikaeYoteiData "
    '//2007/04/19 WAOは金額入力無しなので条件をはずす
    'sql = sql & "       WHERE (NVL(faskgk,0) > 0 OR NVL(fahkgk,0) > 0) "
    sql = sql & "     )"
    sql = sql & "  AND CANWDT IS NULL"
    cnt = gdDBS.Database.ExecuteSQL(sql)
    
    '//累積
    sql = "INSERT INTO tfFurikaeYoteiTran "
    sql = sql & " SELECT * FROM tfFurikaeYoteiData"
    sql = sql & WhereSQL
    cnt = gdDBS.Database.ExecuteSQL(sql)
    '//累積した分を削除
    sql = " DELETE tfFurikaeYoteiData"
    sql = sql & WhereSQL
    Call gdDBS.Database.ExecuteSQL(sql)
'//2003/02/04 次回振込日・次回口座振替日 を更新する
    Dim KouFuriDay As Integer, FurikomiDay As Integer
    Dim KouFuriDate As Date, FurikomiDate As Date

    '//振込日：契約者宛て
    FurikomiDay = gdDBS.SystemUpdate("AAFKDT")
    '//翌月の振込日が算出される
    FurikomiDate = DateSerial( _
                        Mid(gdDBS.SystemUpdate("AANXFK"), 1, 4), _
                        Mid(gdDBS.SystemUpdate("AANXFK"), 5, 2) + 1, _
                        FurikomiDay _
                    )
    '//次回振込日 設定
    gdDBS.SystemUpdate("AANXFK") = Format(NextDay(FurikomiDate), "yyyymmdd")
    
    '//口座振替日：保護者宛て
    KouFuriDay = gdDBS.SystemUpdate("AAKZDT")
    '//翌月の口座振替日が算出される
'//2010/02/23 ２０１０年２月は 2/27,28 が営業日でない為、振替日が 3/1 になってしまっているので１ヶ月先を設定してしまうバグ対応
    Dim wDay As Integer, addMonth As Integer
    wDay = Right(gdDBS.SystemUpdate("AANXKZ"), 2)
    If KouFuriDay <= wDay Then
        addMonth = 1
    End If
    KouFuriDate = DateSerial( _
                        Mid(gdDBS.SystemUpdate("AANXKZ"), 1, 4), _
                        Mid(gdDBS.SystemUpdate("AANXKZ"), 5, 2) + addMonth, _
                        KouFuriDay _
                    )
    KouFuriDate = NextDay(KouFuriDate)
    '//次回口座振替日 設定
    gdDBS.SystemUpdate("AANXKZ") = Format(NextDay(KouFuriDate), "yyyymmdd")

'//2004/04/12 口座振替日を比較して以降の日に 再設定
    If FurikomiDate < KouFuriDate Then      '//年月日
        If FurikomiDay < KouFuriDay Then    '//　　日
            FurikomiDate = DateSerial( _
                                Mid(gdDBS.SystemUpdate("AANXKZ"), 1, 4), _
                                Mid(gdDBS.SystemUpdate("AANXKZ"), 5, 2) + 1, _
                                FurikomiDay _
                            )
        Else
            FurikomiDate = DateSerial( _
                                Mid(gdDBS.SystemUpdate("AANXKZ"), 1, 4), _
                                Mid(gdDBS.SystemUpdate("AANXKZ"), 5, 2) + 0, _
                                FurikomiDay _
                            )
        End If
        '//次回振込日 再設定
        gdDBS.SystemUpdate("AANXFK") = Format(NextDay(FurikomiDate), "yyyymmdd")
    End If
'//2004/05/17 累積件数のログ追加
    Call gdDBS.AutoLogOut(mCaption, "口座振替ＤＢ累積 = " & cnt & " 件 対象 = " & RuisekiDate)

    lblMessage.Caption = mExeMsg & cnt & " 件のデータが累積されました。"
    '//実行更新フラグ設定
    gdDBS.SystemUpdate("AAUPDE") = 1
    Call gdDBS.Database.CommitTrans
    Exit Sub
cmdExec_ClickError:
    Call gdDBS.Database.Rollback
    Call gdDBS.ErrorCheck       '//エラートラップ
'// gdDBS.ErrorCheck() の上に移動
'//    Call gdDBS.Database.Rollback
End Sub

Private Function NextDay(vStartDate As Variant) As Variant
    Dim ix As Integer
    Dim dyn As Object, sql As String
    '//２０連休は無いだろう!!!
    For ix = 0 To 20
        NextDay = DateSerial(Year(vStartDate), Month(vStartDate), Day(vStartDate) + ix)
        '//1=日曜日,2=月曜日...,7=土曜日 なので２以上は月曜日から金曜日のはず
        If (Weekday(NextDay, vbSunday) Mod 7) >= 2 Then
            sql = "SELECT EADATE FROM teHolidayMaster "
            sql = sql & " WHERE EADATE = " & Format(NextDay, "yyyymmdd")
            Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
            If dyn.EOF() Then
                Exit Function
            End If
            Call dyn.Close
        End If
    Next ix
    '//オーバーしたので...。
    NextDay = vStartDate
End Function

Private Sub Form_Load()
    Dim reg As New RegistryClass
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    lblMessage.Caption = mExeMsg
    
    '//ListBox に現在の予定を全てリストアップする。
'    Dim sql As String, dyn As OraDynaset
    Dim sql As String, dyn As Object
    sql = "SELECT FASQNO,TO_CHAR(TO_DATE(FASQNO,'YYYYMMDD'),'YYYY/MM/DD') AS FaDate"
    sql = sql & " FROM tfFurikaeYoteiData"
    sql = sql & " GROUP BY FASQNO"
    sql = sql & " ORDER BY FASQNO"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    Call lstFurikaeBi.Clear
    Do Until dyn.EOF()
        Call lstFurikaeBi.AddItem(dyn.Fields("FaDate"))
'        lstFurikaeBi.Selected(lstFurikaeBi.NewIndex) = True
        Call dyn.MoveNext
    Loop
    Call dyn.Close
#If 0 Then
'''//チェックボックステストのために作成
    Dim i As Integer
    For i = 1 To 10
        Call lstFurikaeBi.AddItem(Format(Now() + i, "yyyy/mm/dd"))
        lstFurikaeBi.Selected(lstFurikaeBi.NewIndex) = True
    Next i
#End If
    cmdExec.Enabled = lstFurikaeBi.ListCount > 0
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mForm = Nothing
    Set frmFurikaeDataRuiseki = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub lstFurikaeBi_ItemCheck(Item As Integer)
    '//チェックボックスは常にチェック状態に維持する
'    lstFurikaeBi.Selected(Item) = True
End Sub

Private Sub mnuEnd_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

