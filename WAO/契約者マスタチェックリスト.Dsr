VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptKeiyakushaCheckList 
   Caption         =   "���������s_WAO - rptKeiyakushaCheckList (ActiveReport)"
   ClientHeight    =   9375
   ClientLeft      =   1545
   ClientTop       =   1890
   ClientWidth     =   15405
   WindowState     =   2  '�ő剻
   _ExtentX        =   27173
   _ExtentY        =   16536
   SectionData     =   "�_��҃}�X�^�`�F�b�N���X�g.dsx":0000
End
Attribute VB_Name = "rptKeiyakushaCheckList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public mYubinCode As String
'Public mYubinName As String
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
    Call mReport.Setup(Me, vbPRPSA4, vbPRORLandscape)
    'txtShoriDate.Text = Mid(gdADO.SystemData("MASRDT"), 1, 4) & " �N " & Mid(gdADO.SystemData("MASRDT"), 5, 2) & " ���x"
    Me.PageSettings.TopMargin = 900
    Me.PageSettings.LeftMargin = 800
    Me.PageSettings.BottomMargin = 300
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
    If "" <> Me.Fields("BAZPC1").Value Then
        If "" <> Me.Fields("BAZPC2").Value Then
            txtBAZPC1.Text = txtBAZPC1.Text & "-" & Me.Fields("BAZPC2").Value
        End If
    End If
    Call pDate_Format(txtBAKYST)
    Call pDate_Format(txtBAKYED)
    Select Case Me.Fields("BAKZSB").Value
    Case 1: txtBAKZSB.Text = "����"
    Case 2: txtBAKZSB.Text = "����"
    End Select
    txtBAADDT.Text = Format(Me.Fields("BAADDT").Value, "yyyy/MM/dd hh:nn:ss")
    txtBAUPDT.Text = Format(Me.Fields("BAUPDT").Value, "yyyy/MM/dd hh:nn:ss")
    mNewCnt(eCount.eItaku) = mNewCnt(eCount.eItaku) + Me.Fields("NEWCNT").Value
    mEdtCnt(eCount.eItaku) = mEdtCnt(eCount.eItaku) + Me.Fields("EDTCNT").Value
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
    lblSysDate.Caption = "( " & Format(Now(), "yyyy/MM/dd hh:nn:ss") & " )"
    lblPage.Caption = "Page: " & Me.PageNumber
End Sub

Private Sub ReportFooter_Format()
    lblTotalMsg.Caption = "<< �V�K�����F " & Format(mNewCnt(eCount.eTotal), "#,#0") & " ��"
    lblTotalMsg.Caption = lblTotalMsg.Caption & " / �ύX�����F " & Format(mEdtCnt(eCount.eTotal), "#,#0") & " ��"
    lblTotalMsg.Caption = lblTotalMsg.Caption & " / �������F " & Format(mNewCnt(eCount.eTotal) + mEdtCnt(eCount.eTotal), "#,#0") & " ��"
    lblTotalMsg.Caption = lblTotalMsg.Caption & " >>"
End Sub
