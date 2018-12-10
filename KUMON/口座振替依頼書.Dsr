VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptKouzaFurikaeIraisho 
   Caption         =   "料金回収代行システム - rptKouzaFurikaeIraisho (ActiveReport)"
   ClientHeight    =   8805
   ClientLeft      =   825
   ClientTop       =   1815
   ClientWidth     =   17895
   WindowState     =   2  '最大化
   _ExtentX        =   31565
   _ExtentY        =   15531
   SectionData     =   "口座振替依頼書.dsx":0000
End
Attribute VB_Name = "rptKouzaFurikaeIraisho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mYubinCode As String
Public mYubinName As String
Public mStartDate As Variant

Private mReport As New ActiveReportClass
Private Enum eCount
    eItaku
    eTotal
End Enum
Private mEdtCnt(0 To 1) As Long
Private mNewCnt(0 To 1) As Long
Private mEdtAndNewCnt(0 To 1) As Long
Private mLineCount As Long

'//履歴のデータで同じ内容を消す操作の為に追加
Private Const cNON_GRP As String = "@@@"
Private mCAKYCD As String
Private mCAKSCD As String
Private mCAHGCD As String

Private Sub ActiveReport_Initialize()
    Call mReport.Setup(Me)
    'txtShoriDate.Text = Mid(gdADO.SystemData("MASRDT"), 1, 4) & " 年 " & Mid(gdADO.SystemData("MASRDT"), 5, 2) & " 月度"
    Me.PageSettings.TopMargin = 500
    Me.PageSettings.LeftMargin = 500
    Me.PageSettings.RightMargin = 500
    Me.PageSettings.BottomMargin = 500
    lblCondition.Caption = ""
    mCAKYCD = cNON_GRP
    mCAKSCD = cNON_GRP
    mCAHGCD = cNON_GRP
    lblHeadKinyuKikanName.Caption = "口座振替" & vbCrLf & "金融機関名称"
End Sub

Private Sub pDate_Format(vData As Object)
    If 0 = vData.DataValue Or vData.DataValue = 20991231 Then
        vData.Text = "---"
    ElseIf vData <> "" Then
        vData.Text = Mid(vData.Text, 1, 4) & "/" & Mid(vData.Text, 5, 2) & "/" & Mid(vData.Text, 7, 2)
    End If
End Sub

Private Sub ActiveReport_ReportStart()
    Erase mNewCnt
    Erase mEdtCnt
    Erase mEdtAndNewCnt
End Sub

Private Sub Detail_BeforePrint()
'//この位置でマスクしないとうまく出来ない
    mLineCount = mLineCount + 1
    'Me.txtCAKSCD = mLineCount
    shpMask.BackStyle = IIf(mLineCount Mod 2, ddBKTransparent, ddBKNormal)
End Sub

Private Sub Detail_Format()
'//微妙にフォーマットタイミングが合わないので良いかぁ!
    '//履歴のデータをグループ化
    If mCAKYCD = Me.Fields("CAKYCD") Then
        txtCAKYCD.Text = ""
        If mCAKSCD = Me.Fields("CAKSCD") Then
            txtCAKSCD.Text = ""
            If mCAHGCD = Me.Fields("CAHGCD") Then
                txtCAHGCD.Text = ""
            End If
        End If
    End If
    mCAKSCD = Me.Fields("CAKSCD")
    mCAKYCD = Me.Fields("CAKYCD")
    mCAHGCD = Me.Fields("CAHGCD")
    
    Call pDate_Format(txtCAKYST)
    Call pDate_Format(txtCAKYED)
    Call pDate_Format(txtCAFKST)
    Call pDate_Format(txtCAFKED)
    If 1 = Me.Fields("rKUBUN") Then
        txtINPDATE.Text = Format(Me.Fields("CAUPDT"), "yyyy/MM/dd HH:nn")
    Else
        txtINPDATE.Text = Format(Me.Fields("CAMKDT"), "yyyy/MM/dd HH:nn")
    End If
    If 0 <> Me.Fields("CAKKBN").Value Then
        txtCABANK.Text = mYubinCode
        txtBANKNAME.Text = mYubinName
	If True = IsNull(Me.Fields("CAYBTK").Value) Then
	        txtCAKZSB.Text = ""
	Else
	        txtCAKZSB.Text = Me.Fields("CAYBTK").Value
	End If
	If True = IsNull(Me.Fields("CAYBTN").Value) Then
        	txtCAKZNO.Text = ""
	Else
        	txtCAKZNO.Text = Me.Fields("CAYBTN").Value
	End If
    End If
    If IsNull(txtCANWDT.DataValue) Then
        txtCANWDT.Text = "新規"
        If Me.Fields("CAUSID") = MainModule.gcImportHogoshaUser Then
            txtCANWDT.Text = txtCANWDT.Text & vbCrLf & "取込"
'        Else
'            txtCANWDT.Text = txtCANWDT.Text & vbCrLf & "入力"
        End If
        mNewCnt(eCount.eItaku) = mNewCnt(eCount.eItaku) + 1
    Else
        txtCANWDT.Text = ""
        If Me.Fields("CAUSID") = MainModule.gcImportHogoshaUser Then
            txtCANWDT.Text = txtCANWDT.Text & vbCrLf & "取込"
'        Else
'            txtCANWDT.Text = txtCANWDT.Text & vbCrLf & "入力"
        End If
        mEdtCnt(eCount.eItaku) = mEdtCnt(eCount.eItaku) + Me.Fields("rKUBUN").Value
    End If
End Sub

Private Sub ItakushaGroupFooter_Format()
    lblGroupMsg.Caption = "< 新規件数： " & Format(mNewCnt(eCount.eItaku), "#,#0") & " 件"
    lblGroupMsg.Caption = lblGroupMsg.Caption & " / 変更件数： " & Format(mEdtCnt(eCount.eItaku), "#,#0") & " 件"
    lblGroupMsg.Caption = lblGroupMsg.Caption & " / 総件数： " & Format(mNewCnt(eCount.eItaku) + mEdtCnt(eCount.eItaku), "#,#0") & " 件"
    lblGroupMsg.Caption = lblGroupMsg.Caption & " >  "
    mNewCnt(eCount.eTotal) = mNewCnt(eCount.eTotal) + mNewCnt(eCount.eItaku)
    mEdtCnt(eCount.eTotal) = mEdtCnt(eCount.eTotal) + mEdtCnt(eCount.eItaku)
    mNewCnt(eCount.eItaku) = 0
    mEdtCnt(eCount.eItaku) = 0
End Sub

Private Sub PageFooter_BeforePrint()
'//この位置でマスクしないとうまく出来ない
    mLineCount = 0
'//微妙にフォーマットタイミングが合わないので良いかぁ!
''    mCAKYCD = cNON_GRP
''    mCAKSCD = cNON_GRP
''    mCAHGCD = cNON_GRP
End Sub

Private Sub PageHeader_Format()
    lblSysDate.Caption = "( " & Format(Now(), "yyyy/mm/dd hh:nn:ss") & " )"
    lblPage.Caption = "Page: " & Me.PageNumber
End Sub

Private Sub ReportFooter_Format()
    lblTotalMsg.Caption = "<< 新規件数： " & Format(mNewCnt(eCount.eTotal), "#,#0") & " 件"
    lblTotalMsg.Caption = lblTotalMsg.Caption & " / 変更件数： " & Format(mEdtCnt(eCount.eTotal), "#,#0") & " 件"
    lblTotalMsg.Caption = lblTotalMsg.Caption & " / 総件数： " & Format(mNewCnt(eCount.eTotal) + mEdtCnt(eCount.eTotal), "#,#0") & " 件"
    lblTotalMsg.Caption = lblTotalMsg.Caption & " >>"
End Sub

