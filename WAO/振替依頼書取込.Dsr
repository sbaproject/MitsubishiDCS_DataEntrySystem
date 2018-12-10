VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptFurikaeReqImport 
   Caption         =   "料金回収代行_WAO - rptFurikaeReqImport (ActiveReport)"
   ClientHeight    =   10050
   ClientLeft      =   720
   ClientTop       =   2100
   ClientWidth     =   16500
   WindowState     =   2  '最大化
   _ExtentX        =   29104
   _ExtentY        =   17727
   SectionData     =   "振替依頼書取込.dsx":0000
End
Attribute VB_Name = "rptFurikaeReqImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#Const NORMAL_OUTPUT = False

Public mTotalCnt As Long    '//呼び出しフォームでセットされる

Private mReport As New ActiveReportClass
Private mRimp As New FurikaeReqImpClass
Private Enum eCount
'//2006/04/05 正常データは印刷しない
#If NORMAL_OUTPUT = True Then
    eNoMstUpd
#End If
    eInvalid
    eWarning
#If NORMAL_OUTPUT = True Then
    eTotal
#End If
End Enum
#If NORMAL_OUTPUT = True Then
Private mDataCnt(0 To 3) As Long
#Else
Private mDataCnt(0 To 1) As Long
#End If
Private mLineCount As Long

Private Sub ActiveReport_Initialize()
    Call mReport.Setup(Me)
    'txtShoriDate.Text = Mid(gdADO.SystemData("MASRDT"), 1, 4) & " 年 " & Mid(gdADO.SystemData("MASRDT"), 5, 2) & " 月度"
    Me.PageSettings.TopMargin = 500
    Me.PageSettings.LeftMargin = 500
    Me.PageSettings.RightMargin = 500
    Me.PageSettings.BottomMargin = 500
    '//この時点は Load()
    'mDataCnt(eCount.eTotal) = mTotalCnt
End Sub

Private Sub ActiveReport_ReportStart()
    Erase mDataCnt
    '//ここでしないと取れない
#If NORMAL_OUTPUT = True Then
    mDataCnt(eCount.eTotal) = mTotalCnt
#End If
End Sub

Private Sub Detail_BeforePrint()
'//この位置でマスクしないとうまく出来ない
    mLineCount = mLineCount + 1
    shpMask.BackStyle = IIf(mLineCount Mod 2, ddBKTransparent, ddBKNormal)
End Sub

Private Sub pTextBoxColor(vObj As Object, vStatus As Variant)
    Select Case vStatus
    Case mRimp.errInvalid
        vObj.ForeColor = vbRed
    Case mRimp.errWarning
        vObj.ForeColor = vbMagenta
    Case mRimp.errEditData
        vObj.ForeColor = vbGreen
    Case Else
        vObj.ForeColor = vbBlack
    End Select
    vObj.Font.Underline = vStatus <> 0
End Sub

Private Sub Detail_Format()
    Select Case Me.Fields("CiERROr")
    Case mRimp.errInvalid, mRimp.errEditData, mRimp.errImport
        mDataCnt(eCount.eInvalid) = mDataCnt(eCount.eInvalid) + 1
    Case mRimp.errWarning
        mDataCnt(eCount.eWarning) = mDataCnt(eCount.eWarning) + 1
    Case Else
    End Select
    If Not IsNull(Me.Fields("CiKKBN").Value) Then
        '//振替先が郵便局なら口座種別に通帳記号を...。
        If "郵便局" = Me.Fields("CiKKBN").Value Then
            txtCIKZSB.Text = Me.Fields("CiYBTK").Value
            txtCIKZNO.Text = Me.Fields("CiYBTN").Value
        End If
    End If
    If Not IsNull(Me.Fields("CiFKST").Value) Then
        txtCIFKST.Text = Format(CDate(Me.Fields("CiFKST").Value), "yyyy/MM")
    End If
    Call pTextBoxColor(txtCIERRNM, Me.Fields("CiERROr").Value)
    Call pTextBoxColor(txtABKJNM, Me.Fields("CiITKBe").Value)
    Call pTextBoxColor(txtCiKYCD, Me.Fields("CiKYCDe").Value)
    'Call pTextBoxColor(txtCIKSCD, Me.Fields("CiKSCDe").Value)
    Call pTextBoxColor(txtCiHGCD, Me.Fields("CiHGCDe").Value)
    Call pTextBoxColor(txtCIKJNM, Me.Fields("CiKJNMe").Value)
    Call pTextBoxColor(txtCIKNNM, Me.Fields("CiKNNMe").Value)
    Call pTextBoxColor(txtCiSTNM, Me.Fields("CiSTNMe").Value)
    Call pTextBoxColor(txtCIKKBN, Me.Fields("CiKKBNe").Value)
    Call pTextBoxColor(txtCiBANK, Me.Fields("CiBANKe").Value)
    Call pTextBoxColor(txtCISITN, Me.Fields("CiSITNe").Value)
    Call pTextBoxColor(txtCIBKNM, Me.Fields("CiBKNMe").Value)
    Call pTextBoxColor(txtCISINM, Me.Fields("CiSINMe").Value)
    Call pTextBoxColor(txtDABKNM, Me.Fields("CiBKNMe").Value)
    Call pTextBoxColor(txtDASTNM, Me.Fields("CiSTNMe").Value)
    Call pTextBoxColor(txtCIKZSB, Me.Fields("CiKZSBe").Value)
    Call pTextBoxColor(txtCIKZNO, Me.Fields("CiKZNOe").Value)
    'Call pTextBoxColor(txtCIKZNM, Me.Fields("CiKZNMe").Value)
    'Call pTextBoxColor(txtCiSKGK, Me.Fields("CiSKGKe").Value)
    Call pTextBoxColor(txtCIFKST, Me.Fields("CiFKSTe").Value)

'//2006/04/05 正常データは印刷しない
#If NORMAL_OUTPUT = True Then
    '//2006/04/04 マスタ反映ＯＫフラグ追加
    If 0 <> Val(Me.Fields("CiMUPD").Value) Then
        txtCIERRNM.Text = "正常 => ×反映しない"
        txtCIERRNM.Font.Underline = True
        mDataCnt(eCount.eNoMstUpd) = mDataCnt(eCount.eNoMstUpd) + 1
    End If
#End If
End Sub

Private Sub PageHeader_Format()
    lblSysDate.Caption = "( " & Format(Now(), "yyyy/mm/dd hh:nn:ss") & " )"
    lblPage.Caption = "Page: " & Me.PageNumber
End Sub

Private Sub ReportFooter_Format()
    '//正常データは出力しないので加減算して算出
    lblTotalMsg.Caption = ""
    lblTotalMsg.Caption = lblTotalMsg.Caption & " 異常： " & Format(mDataCnt(eCount.eInvalid), "#,#0") & " 件 / "
    lblTotalMsg.Caption = lblTotalMsg.Caption & " 警告： " & Format(mDataCnt(eCount.eWarning), "#,#0") & " 件 / "
'//2006/04/05 正常データは印刷しない
#If NORMAL_OUTPUT = True Then
    lblTotalMsg.Caption = lblTotalMsg.Caption & " 正常： " & Format(mDataCnt(eCount.eTotal) - (mDataCnt(eCount.eInvalid) + mDataCnt(eCount.eWarning) + mDataCnt(eCount.eNoMstUpd)), "#,#0") & " 件 / "
    lblTotalMsg.Caption = lblTotalMsg.Caption & " 除外： " & Format(mDataCnt(eCount.eNoMstUpd), "#,#0") & " 件 / "
    lblTotalMsg.Caption = lblTotalMsg.Caption & " 総件数： " & Format(mDataCnt(eCount.eTotal), "#,#0") & " 件"
#Else
    lblTotalMsg.Caption = lblTotalMsg.Caption & " 総件数： " & Format(mDataCnt(eCount.eInvalid) + mDataCnt(eCount.eWarning), "#,#0") & " 件"
#End If
End Sub

