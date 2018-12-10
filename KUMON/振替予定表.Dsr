VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptYoteiDataCheckList 
   Caption         =   "料金回収代行システム - rptYoteiDataCheckList (ActiveReport)"
   ClientHeight    =   11145
   ClientLeft      =   345
   ClientTop       =   2265
   ClientWidth     =   17895
   WindowState     =   2  '最大化
   _ExtentX        =   31565
   _ExtentY        =   19659
   SectionData     =   "振替予定表.dsx":0000
End
Attribute VB_Name = "rptYoteiDataCheckList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mYubinCode As String
Public mYubinName As String
Public mStartDate As Variant

Private mReport As New ActiveReportClass
Private Enum eSum
    eKeiyaku
    eItaku
    eTotal
End Enum
Private mEdtCnt(0 To 2) As Long
Private mEdtGaku(0 To 2) As Currency
Private mNewCnt(0 To 2) As Long
Private mKaiyaku(0 To 2) As Long

Private mLineCount As Long

Private Sub ActiveReport_Initialize()
    Call mReport.Setup(Me, vOrientation:=vbPRORPortrait)
    'txtShoriDate.Text = Mid(gdADO.SystemData("MASRDT"), 1, 4) & " 年 " & Mid(gdADO.SystemData("MASRDT"), 5, 2) & " 月度"
    Me.PageSettings.TopMargin = 300
    Me.PageSettings.LeftMargin = 300
    Me.PageSettings.RightMargin = 300
    Me.PageSettings.BottomMargin = 300
    lblCondition.Caption = ""
    mYubinCode = gdDBS.SystemUpdate("AAYSNO")
End Sub

Private Sub pDate_Format(vData As Object)
    If 0 = vData.DataValue Or vData.DataValue = 20991231 Then
        vData.Text = "---"
    Else
        vData.Text = Mid(vData.Text, 1, 4) & "/" & Mid(vData.Text, 5, 2) & "/" & Mid(vData.Text, 7, 2)
    End If
End Sub

Private Sub ActiveReport_ReportStart()
    Erase mEdtCnt
    Erase mEdtGaku
    Erase mNewCnt
    Erase mKaiyaku
End Sub

Private Sub Detail_BeforePrint()
'//この位置でマスクしないとうまく出来ない
    If mLineCount Then
        txtCAKYCD.Text = ""
        txtCAKSCD.Text = ""
    End If
    mLineCount = mLineCount + 1
    shpMask.BackStyle = IIf(mLineCount Mod 2, ddBKTransparent, ddBKNormal)
End Sub

Private Sub Detail_Format()
    '//2004/07/15 金融機関が変更されている場合のみ変更ざれた金融機関の内容を出力
    txtEditKouzaData.Text = ""
    If Me.Fields("CAKKBN").Value <> Me.Fields("FAKKBN").Value Then
        If Me.Fields("CAKKBN").Value = eBankKubun.YuubinKyoku Then
            txtEditKouzaData.Text = mYubinCode & "  " & Me.Fields("CAYBTK").Value & "  " & Me.Fields("CAYBTN").Value
        ElseIf Me.Fields("CAKKBN").Value = eBankKubun.KinnyuuKikan Then
            txtEditKouzaData.Text = Me.Fields("CABANK").Value & "  " & Me.Fields("CASITN").Value & "  " & Me.Fields("CAKZSB").Value & "  " & Me.Fields("CAKZNO").Value
        End If
    ElseIf Me.Fields("CAKKBN").Value = eBankKubun.YuubinKyoku Then
        If Me.Fields("CAYBTK").Value <> Me.Fields("FAYBTK").Value _
        Or Me.Fields("CAYBTN").Value <> Me.Fields("FAYBTN").Value Then
            txtEditKouzaData.Text = mYubinCode & "  " & Me.Fields("CAYBTK").Value & "  " & Me.Fields("CAYBTN").Value
        End If
    ElseIf Me.Fields("CAKKBN").Value = eBankKubun.KinnyuuKikan Then
        If Me.Fields("CABANK").Value <> Me.Fields("FABANK").Value _
        Or Me.Fields("CASITN").Value <> Me.Fields("FASITN").Value _
        Or Me.Fields("CAKZSB").Value <> Me.Fields("FAKZSB").Value _
        Or Me.Fields("CAKZNO").Value <> Me.Fields("FAKZNO").Value Then
            txtEditKouzaData.Text = Me.Fields("CABANK").Value & "  " & Me.Fields("CASITN").Value & "  " & Me.Fields("CAKZSB").Value & "  " & Me.Fields("CAKZNO").Value
        End If
    End If
    
    Call pDate_Format(txtCAKYST)
    Call pDate_Format(txtCAKYED)
    Call pDate_Format(txtCAFKST)
    Call pDate_Format(txtCAFKED)
    If IsNull(txtCANWDT.DataValue) Then
        txtCANWDT.Text = "新規"
        mNewCnt(eSum.eKeiyaku) = mNewCnt(eSum.eKeiyaku) + 1
    Else
        txtCANWDT.Text = ""
    End If
    mEdtCnt(eSum.eKeiyaku) = mEdtCnt(eSum.eKeiyaku) + 1
    mEdtGaku(eSum.eKeiyaku) = mEdtGaku(eSum.eKeiyaku) + gdDBS.Nz(Me.Fields("CASKGK").Value, 0)  '//保護者マスタ内データ(CA ?)が変更データ
    '//請求金額＝変更後金額の場合
    If txtFASKGK.DataValue = txtCASKGK.DataValue Then
        txtCASKGK.Text = "---"
    End If
    mKaiyaku(eSum.eKeiyaku) = mKaiyaku(eSum.eKeiyaku) + Abs(Sgn(gdDBS.Nz(Me.Fields("CAKYFG").Value, 0)))
End Sub

Private Sub ItakushaGroupFooter_Format()
    If mKaiyaku(eSum.eItaku) Then
        txtKaiyakuTtlCnt.Text = "解約件数  " & Format(mKaiyaku(eSum.eItaku), "#,##0") & " 件"
    Else
        txtKaiyakuTtlCnt.Text = ""
    End If
    If mNewCnt(eSum.eItaku) Then
        txtNewDataTtlCnt.Text = "新規件数  " & Format(mNewCnt(eSum.eItaku), "#,##0") & " 件"
    Else
        txtNewDataTtlCnt.Text = ""
    End If
    mNewCnt(eSum.eTotal) = mNewCnt(eSum.eTotal) + mNewCnt(eSum.eItaku)
    mEdtCnt(eSum.eTotal) = mEdtCnt(eSum.eTotal) + mEdtCnt(eSum.eItaku)
    mEdtGaku(eSum.eTotal) = mEdtGaku(eSum.eTotal) + mEdtGaku(eSum.eItaku)
    mKaiyaku(eSum.eTotal) = mKaiyaku(eSum.eTotal) + mKaiyaku(eSum.eItaku)
    mNewCnt(eSum.eItaku) = 0
    mEdtCnt(eSum.eItaku) = 0
    mEdtGaku(eSum.eItaku) = 0
    mKaiyaku(eSum.eItaku) = 0
End Sub

Private Sub KeiyakushaGroupFooter_Format()
    If mEdtCnt(eSum.eKeiyaku) Then
        txtHenkoSubCnt.Text = Format(mEdtCnt(eSum.eKeiyaku), "#,##0") & " 件"
    Else
        txtHenkoSubCnt.Text = ""
    End If
    If mEdtGaku(eSum.eKeiyaku) Then
        txtHenkoSubGaku.Text = Format(mEdtGaku(eSum.eKeiyaku), "#,##0")
    Else
        txtHenkoSubGaku.Text = ""
    End If
    If mKaiyaku(eSum.eKeiyaku) Then
        txtKaiyakuSubCnt.Text = Format(mKaiyaku(eSum.eKeiyaku), "#,##0") & " 件"
    Else
        txtKaiyakuSubCnt.Text = ""
    End If
    If mNewCnt(eSum.eKeiyaku) Then
        txtNewDataSubCnt.Text = Format(mNewCnt(eSum.eKeiyaku), "#,##0") & " 件"
    Else
        txtNewDataSubCnt.Text = ""
    End If
    mNewCnt(eSum.eItaku) = mNewCnt(eSum.eItaku) + mNewCnt(eSum.eKeiyaku)
    mEdtCnt(eSum.eItaku) = mEdtCnt(eSum.eItaku) + mEdtCnt(eSum.eKeiyaku)
    mEdtGaku(eSum.eItaku) = mEdtGaku(eSum.eItaku) + mEdtGaku(eSum.eKeiyaku)
    mKaiyaku(eSum.eItaku) = mKaiyaku(eSum.eItaku) + mKaiyaku(eSum.eKeiyaku)
    mNewCnt(eSum.eKeiyaku) = 0
    mEdtCnt(eSum.eKeiyaku) = 0
    mEdtGaku(eSum.eKeiyaku) = 0
    mKaiyaku(eSum.eKeiyaku) = 0
    mLineCount = 0
End Sub

Private Sub PageFooter_BeforePrint()
'//この位置でマスクしないとうまく出来ない
    mLineCount = 0
End Sub

Private Sub PageHeader_Format()
    lblSysDate.Caption = "( " & Format(Now(), "yyyy/mm/dd hh:nn:ss") & " )"
    lblPage.Caption = "Page: " & Me.PageNumber
    Me.txtFurikaeDate.Text = gdDBS.SystemUpdate("AANXKZ")
    Me.txtFurikaeDate.Text = Mid(Me.txtFurikaeDate.Text, 1, 4) & "/" & Mid(Me.txtFurikaeDate.Text, 5, 2) & "/" & Mid(Me.txtFurikaeDate.Text, 7, 2)
    Me.txtFurikomiDate.Text = gdDBS.SystemUpdate("AANXFK")
    Me.txtFurikomiDate.Text = Mid(Me.txtFurikomiDate.Text, 1, 4) & "/" & Mid(Me.txtFurikomiDate.Text, 5, 2) & "/" & Mid(Me.txtFurikomiDate.Text, 7, 2)
End Sub

