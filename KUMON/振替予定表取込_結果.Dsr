VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptYoteiReqImportReport 
   Caption         =   "料金回収代行システム - rptYoteiReqImportReport (ActiveReport)"
   ClientHeight    =   11145
   ClientLeft      =   1815
   ClientTop       =   1290
   ClientWidth     =   16080
   WindowState     =   2  '最大化
   _ExtentX        =   28363
   _ExtentY        =   19659
   SectionData     =   "振替予定表取込_結果.dsx":0000
End
Attribute VB_Name = "rptYoteiReqImportReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mReport As New ActiveReportClass

Private mLineCount As Long
Private mGrpItakuCode As String

Private Sub ActiveReport_Initialize()
    Call mReport.Setup(Me, vOrientation:=vbPRORPortrait)
    'txtShoriDate.Text = Mid(gdADO.SystemData("MASRDT"), 1, 4) & " 年 " & Mid(gdADO.SystemData("MASRDT"), 5, 2) & " 月度"
    Me.PageSettings.TopMargin = 500
    Me.PageSettings.LeftMargin = 500
    Me.PageSettings.RightMargin = 500
    Me.PageSettings.BottomMargin = 500
    lblCondition.Caption = ""
    mGrpItakuCode = "@@@"
    'lblLeftTtGaku.Caption = "変更後 " & vbCrLf & "合計金額"
    'lblRightTtGaku.Caption = lblLeftTtGaku.Caption
End Sub

Private Sub BadDataFormat(vMeisai As DDActiveReports2.Field, vTotal As DDActiveReports2.Field)
    vMeisai.Font.Underline = Me.Fields(vMeisai.DataField) <> Me.Fields(vTotal.DataField)
    'vMeisai.Font.Bold = Me.Fields(vMeisai.DataField) <> Me.Fields(vTotal.DataField)
    If vMeisai.Font.Underline Then
        vMeisai.ForeColor = vbRed
    Else
        vMeisai.ForeColor = vbBlack
    End If
    vTotal.Font.Underline = vMeisai.Font.Underline
    'vTotal.Font.Bold = vMeisai.Font.Bold
    vTotal.ForeColor = vMeisai.ForeColor
End Sub

Private Sub Detail_Format()
    '//変更件数違い
    Call BadDataFormat(txtFIHKCT, txtFIHKCT_T) '.DataField)
    '//変更金額違い
    Call BadDataFormat(txtFIHKKG, txtFIHKKG_T) '.DataField)
    '//解約件数違い
    Call BadDataFormat(txtFIKYCT, txtFIKYCT_T) '.DataField)
End Sub

Private Sub ImportGroupHeader_Format()
    Select Case Me.Fields("FIADID").Value
    Case MainModule.gcYoteiImportToMaster
        txtFIADID.Text = "【マスタ反映】"
    Case MainModule.gcYoteiImportToDelete
        txtFIADID.Text = "【廃棄データ】"
    Case Else
    End Select
    txtFIMKDT.Text = "【" & Me.Fields("FIMKDT").Value & "】"
End Sub

Private Sub Detail_BeforePrint()
'//この位置でマスクしないとうまく出来ない
    mLineCount = mLineCount + 1
    'shpMask.BackStyle = IIf(mLineCount Mod 2, ddBKTransparent, ddBKNormal)
    shpMask.BackStyle = IIf(mLineCount Mod 2, ddBKTransparent, ddBKNormal)
    txtABITCD.Visible = mGrpItakuCode <> Me.Fields("ABITCD")
    mGrpItakuCode = Me.Fields("ABITCD")
End Sub

