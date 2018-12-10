VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptYoteiReqImport 
   Caption         =   "料金回収代行システム - rptYoteiReqImport (ActiveReport)"
   ClientHeight    =   9105
   ClientLeft      =   720
   ClientTop       =   2130
   ClientWidth     =   17490
   WindowState     =   2  '最大化
   _ExtentX        =   30850
   _ExtentY        =   16060
   SectionData     =   "振替予定表取込.dsx":0000
End
Attribute VB_Name = "rptYoteiReqImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mTotalCnt As Long    '//呼び出しフォームでセットされる

Private mReport As New ActiveReportClass
Private mYimp As New FurikaeSchImpClass
Private mGroup As String
Private Enum eCount
    eInvalid
    eWarning
    eTotal
End Enum

Private mBodyEdit   As Integer      '//ボディの変更件数
Private mBodyGaku   As Currency     '//ボディの変更金額
Private mBodyCancel As Integer      '//ボディの解約件数

Private mDataCnt(0 To 2) As Long
Private mLineCount As Long

Private Sub ActiveReport_Initialize()
    '//A4 縦置き
'//2007/06/08 デフォルト用紙をＡ４=>Ｂ４に変更：Ａ４をわざわざＢ４に変更して出力している為
    Call mReport.Setup(Me, vbPRPSB4, ddOPortrait)
    'txtShoriDate.Text = Mid(gdADO.SystemData("MASRDT"), 1, 4) & " 年 " & Mid(gdADO.SystemData("MASRDT"), 5, 2) & " 月度"
    Me.PageSettings.TopMargin = 500
    Me.PageSettings.LeftMargin = 500
    Me.PageSettings.RightMargin = 500
    Me.PageSettings.BottomMargin = 500
    '//この時点は Load()
    'mDataCnt(eCount.eTotal) = mTotalCnt
    mBodyEdit = 0      '//ボディの変更件数
    mBodyGaku = 0      '//ボディの変更金額
    mBodyCancel = 0    '//ボディの解約件数
End Sub

Private Sub ActiveReport_ReportStart()
    '//ここでしないと取れない
'    mDataCnt(eCount.eTotal) = mTotalCnt
    Erase mDataCnt
End Sub

Private Sub Detail_BeforePrint()
'//この位置でマスクしないとうまく出来ない
    mLineCount = mLineCount + 1
    'shpMask.BackStyle = IIf(mLineCount Mod 2, ddBKTransparent, ddBKNormal)
End Sub

Private Sub pTextBoxColor(vObj As Object, vStatus As Variant)
    Select Case vStatus
    Case mYimp.errInvalid
        vObj.ForeColor = vbRed
    Case mYimp.errWarning
        vObj.ForeColor = vbMagenta
    Case mYimp.errEditData
        vObj.ForeColor = vbGreen
    Case Else
        vObj.ForeColor = vbBlack
    End Select
    vObj.Font.Underline = vStatus <> 0
End Sub

Private Sub Detail_Format()
    If Me.Fields("FIRKBN") = mYimp.RecordIsTotal Then
        Me.txtFIHKCT_BODY.Text = mBodyEdit & " 件"      '//ボディの変更件数 計算結果
        Me.txtFIHKKG.Text = Format(mBodyGaku, "#,##0")  '//ボディの変更金額 計算結果
        Me.txtFIKYFG.Text = mBodyCancel & " 件"         '//ボディの解約件数 計算結果
        Me.txtFIHKCT_TOTAL.Text = Me.txtFIHKCT_TOTAL.Text & " 件"
        Me.txtFIKYCT_TOTAL.Text = Me.txtFIKYCT_TOTAL.Text & " 件"
        mBodyEdit = 0      '//ボディの変更件数
        mBodyGaku = 0      '//ボディの変更金額
        mBodyCancel = 0    '//ボディの解約件数
        shpMask.Visible = True
    Else
        '//合計欄の表示を抑止
        Me.txtFIHKCT_BODY.Text = ""
        Me.txtFIHKCT_TOTAL.Text = ""
        Me.txtFIHKKG_TOTAL.Text = ""
        Me.txtFIKYCT_TOTAL.Text = ""
        If Not IsNull(Me.Fields("FIHKKG")) Then '//ボディの変更金額
            mBodyGaku = mBodyGaku + gdDBS.Nz(Me.Fields("FIHKKG")) '//ボディの変更金額
        End If
        If 0 <> gdDBS.Nz(Me.Fields("FIKYFG"), 0) Then
            mBodyCancel = mBodyCancel + 1           '//ボディの解約件数
        Else
            mBodyEdit = mBodyEdit + 1                   '//ボディの変更件数
        End If
        shpMask.Visible = False
    End If
    '//sql = sql & "FIITKB || FIKYCD || FIKSCD || FIPGNO || FIFKDT FIGROUP" & vbCrLf
    If mGroup = Me.Fields("FIGROUP").Value Then
        '//2006/04/05 それぞれにエラーが無ければグルーピングの為に消去
        If Me.Fields("FiITKBe").Value = mYimp.errNormal Then
            Me.txtABITCD.Text = ""
        End If
        If Me.Fields("FiKYCDe").Value = mYimp.errNormal Then
            Me.txtFIKYCD.Text = ""
        End If
        If Me.Fields("FiKSCDe").Value = mYimp.errNormal Then
            Me.txtFIKSCD.Text = ""
        End If
        If Me.Fields("FiPGNOe").Value = mYimp.errNormal Then
            Me.txtFIPGNO.Text = ""
        End If
        If Me.Fields("FiFKDTe").Value = mYimp.errNormal Then
            Me.txtFIFKDT.Text = ""
        End If
    Else
        mGroup = Me.Fields("FIGROUP").Value
    End If
    '//各件数を計算
    mDataCnt(eCount.eTotal) = mDataCnt(eCount.eTotal) + 1
    Select Case Me.Fields("FiERROr")
    Case mYimp.errInvalid, mYimp.errEditData, mYimp.errImport
        mDataCnt(eCount.eInvalid) = mDataCnt(eCount.eInvalid) + 1
    Case mYimp.errWarning
        mDataCnt(eCount.eWarning) = mDataCnt(eCount.eWarning) + 1
    End Select
    Call pTextBoxColor(txtFIERRNM, Me.Fields("FiERROr").Value)
    Call pTextBoxColor(txtFIFKDT, Me.Fields("FiFKDTe").Value)
    Call pTextBoxColor(txtABITCD, Me.Fields("FiITKBe").Value)
    Call pTextBoxColor(txtFIKYCD, Me.Fields("FiKYCDe").Value)
    Call pTextBoxColor(txtFIKSCD, Me.Fields("FiKSCDe").Value)
    Call pTextBoxColor(txtFIPGNO, Me.Fields("FiPGNOe").Value)
    Call pTextBoxColor(txtFIHGCD, Me.Fields("FiHGCDe").Value)
    Call pTextBoxColor(txtFIHKKG, Me.Fields("FiHKKGe").Value)
    Call pTextBoxColor(txtFIKYFG, Me.Fields("FiKYFGe").Value)
    If Me.Fields("FiHKCTe").Value Then
        Call pTextBoxColor(txtFIHKCT_BODY, Me.Fields("FiHKCTe").Value)
    End If
    Call pTextBoxColor(txtFIHKCT_TOTAL, Me.Fields("FiHKCTe").Value)
    Call pTextBoxColor(txtFIHKKG_TOTAL, Me.Fields("FiHKKGe").Value)
    If Me.Fields("FiHKCTe").Value Then
        Call pTextBoxColor(txtFIKYFG, Me.Fields("FiKYCTe").Value)
    End If
    Call pTextBoxColor(txtFIKYCT_TOTAL, Me.Fields("FiKYCTe").Value)
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
    lblTotalMsg.Caption = lblTotalMsg.Caption & " 正常： " & Format(mDataCnt(eCount.eTotal) - (mDataCnt(eCount.eInvalid) + mDataCnt(eCount.eWarning)), "#,#0") & " 件 / "
    lblTotalMsg.Caption = lblTotalMsg.Caption & " 総件数： " & Format(mDataCnt(eCount.eTotal), "#,#0") & " 件"
End Sub

