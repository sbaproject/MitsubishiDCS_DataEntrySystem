VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptKouzaFurikaeIraisho 
   Caption         =   "���������s_WAO - rptKouzaFurikaeIraisho (ActiveReport)"
   ClientHeight    =   8490
   ClientLeft      =   1380
   ClientTop       =   2220
   ClientWidth     =   21735
   WindowState     =   2  '�ő剻
   _ExtentX        =   39714
   _ExtentY        =   14975
   SectionData     =   "�����U�ֈ˗���.dsx":0000
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

Private Sub ActiveReport_Initialize()
    Call mReport.Setup(Me)
    'txtShoriDate.Text = Mid(gdADO.SystemData("MASRDT"), 1, 4) & " �N " & Mid(gdADO.SystemData("MASRDT"), 5, 2) & " ���x"
    Me.PageSettings.TopMargin = 900
    Me.PageSettings.LeftMargin = 500
    Me.PageSettings.BottomMargin = 300
    Me.Printer.PaperSize = vbPRPSB4
    Me.PageSettings.PaperSize = vbPRPSB4
    lblCondition.Caption = ""
'    Dim obj As Object
'    For Each obj In Me.Detail.Controls
'        If UCase(Left(obj.Name, 3)) = UCase("txt") Then
'            obj.BackStyle = 1
'            obj.BackColor = vbRed
'        End If
'    Next obj
    mNewCnt(eCount.eTotal) = 0
    mEdtCnt(eCount.eTotal) = 0
End Sub

Private Sub pDate_Format(vData As Object)
    If 0 = vData.DataValue Or vData.DataValue = 20991231 Then
        vData.Text = "---"
    Else
        vData.Text = Mid(vData.Text, 1, 4) & "/" & Mid(vData.Text, 5, 2) & "/" & Mid(vData.Text, 7, 2)
    End If
End Sub

Private Sub ActiveReport_ReportStart()
    mNewCnt(eCount.eTotal) = 0
    mEdtCnt(eCount.eTotal) = 0
End Sub

Private Sub Detail_BeforePrint()
'//���̈ʒu�Ń}�X�N���Ȃ��Ƃ��܂��o���Ȃ�
    mLineCount = mLineCount + 1
    'Me.txtCAKSCD = mLineCount
    shpMask.BackStyle = IIf(mLineCount Mod 2, ddBKTransparent, ddBKNormal)
End Sub

Private Sub Detail_Format()
    Call pDate_Format(txtCAFKST)
    Call pDate_Format(txtCAFKED)
    If 0 <> Me.Fields("CAKKBN").Value Then
        txtCABANK.Text = mYubinCode
        txtBANKNAME.Text = mYubinName
        txtCAKZSB.Text = Me.Fields("CAYBTK").Value
        txtCAKZNO.Text = Me.Fields("CAYBTN").Value
    End If
    If IsNull(txtCANWDT.DataValue) Then
        txtCANWDT.Text = "�V�K"
        mNewCnt(eCount.eItaku) = mNewCnt(eCount.eItaku) + 1
    Else
        txtCANWDT.Text = ""
        mEdtCnt(eCount.eItaku) = mEdtCnt(eCount.eItaku) + 1
    End If
End Sub

Private Sub ItakushaGroupFooter_Format()
    lblGroupMsg.Caption = "< �V�K�����F " & Format(mNewCnt(eCount.eItaku), "#,#0") & " ��"
    lblGroupMsg.Caption = lblGroupMsg.Caption & " / �ύX�����F " & Format(mEdtCnt(eCount.eItaku), "#,#0") & " ��"
    lblGroupMsg.Caption = lblGroupMsg.Caption & " / �������F " & Format(mNewCnt(eCount.eItaku) + mEdtCnt(eCount.eItaku), "#,#0") & " ��"
    lblGroupMsg.Caption = lblGroupMsg.Caption & " >  "
    mNewCnt(eCount.eTotal) = mNewCnt(eCount.eTotal) + mNewCnt(eCount.eItaku)
    mEdtCnt(eCount.eTotal) = mEdtCnt(eCount.eTotal) + mEdtCnt(eCount.eItaku)
    mNewCnt(eCount.eItaku) = 0
    mEdtCnt(eCount.eItaku) = 0
End Sub

Private Sub PageFooter_BeforePrint()
'//���̈ʒu�Ń}�X�N���Ȃ��Ƃ��܂��o���Ȃ�
    mLineCount = 0
End Sub

Private Sub PageHeader_Format()
    lblSysDate.Caption = "( " & Format(Now(), "yyyy/mm/dd hh:mm:ss") & " )"
    lblPage.Caption = "Page: " & Me.PageNumber
End Sub

Private Sub ReportFooter_Format()
    lblTotalMsg.Caption = "<< �V�K�����F " & Format(mNewCnt(eCount.eTotal), "#,#0") & " ��"
    lblTotalMsg.Caption = lblTotalMsg.Caption & " / �ύX�����F " & Format(mEdtCnt(eCount.eTotal), "#,#0") & " ��"
    lblTotalMsg.Caption = lblTotalMsg.Caption & " / �������F " & Format(mNewCnt(eCount.eTotal) + mEdtCnt(eCount.eTotal), "#,#0") & " ��"
    lblTotalMsg.Caption = lblTotalMsg.Caption & " >>"
End Sub
