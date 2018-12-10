VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActiveReportClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mReport As ActiveReport
Attribute mReport.VB_VarHelpID = -1

#Const ARREPORT_VERSION = "2.0"

Private Enum eToolbarID
    Settings = 1     '����ݒ�

#If ARREPORT_VERSION = "1.5" Then
    Contents = 0        '0   5110    �ڎ�
                        '1   500 -
    PrintOut = 2        '2   5113    ���
                        '3   500 -
    Reduce = 4          '4   5108    �k��
    Expand = 5          '5   5109    �g��
    ZoomRatio = 6       '6   5114    �{��
                        '7   500 -
    BeforePage = 8      '8   5106    �O�y�[�W
    NextPage = 9        '9   5107    ���y�[�W
    GotoPage = 10       '10  5115    �y�[�W
                        '11  500 -
    Back = 12           '12  5111    �߂�
    Forward = 13        '13  5112    �i��
#ElseIf ARREPORT_VERSION = "2.0" Then
                        '�ȉ��ͤ V2.0 ��`�ς݂̃c�[��ID�ꗗ�ł��
    Contents = 0        ' 0 TOC�i���o���ꗗ�j  9c48h
                        ' 1 ------------
    PrintOut = 2        ' 2 ���               8005h
                        ' 3 ------------
    Copy = 4            ' 4 �R�s�[             9c4ch
                        ' 5 ------------
    Search = 6          ' 6 ����               8006h
                        ' 7 ------------
    PageOne = 8         ' 8 �P��y�[�W         9c4Eh
    PageMulti = 9       ' 9 �����y�[�W         9c4Dh
                        '10 ------------
    Reduce = 11         '11 �k��               9c46h
    Expand = 12         '12 �g��               8007h
    ZoomRatio = 13      '13 �{��               8003h
                        '14 ------------
    BeforePage = 15     '15 �O�y�[�W           9c44h
    NextPage = 16       '16 ���y�[�W           9c45h
    GotoPage = 17       '17 �y�[�W�ԍ�         8004h
                        '18 ------------
    Back = 19           '19 �߂�               8008h
    Forward = 20        '20 �i��               9c4Ah
#End If
End Enum

Private Enum eToolbar
    Settings = 1001     '����ݒ�

#If ARREPORT_VERSION = "1.5" Then
    Contents = 5110     '0   5110    �ڎ�
                        '1   500 -
    PrintOut = 5113     '2   5113    ���
                        '3   500 -
    Reduce = 5108       '4   5108    �k��
    Expand = 5109       '5   5109    �g��
    ZoomRatio = 5114    '6   5114    �{��
                        '7   500 -
    BeforePage = 5106   '8   5106    �O�y�[�W
    NextPage = 5107     '9   5107    ���y�[�W
    GotoPage = 5115     '10  5115    �y�[�W
                        '11  500 -
    Back = 5111         '12  5111    �߂�
    Forward = 5112      '13  5112    �i��
#ElseIf ARREPORT_VERSION = "2.0" Then
                        '�ȉ��ͤ V2.0 ��`�ς݂̃c�[��ID�ꗗ�ł��
    Contents = &H9C48   ' 0 TOC�i���o���ꗗ�j  9c48h
                        ' 1 ------------
    PrintOut = &H8005   ' 2 ���               8005h
                        ' 3 ------------
    Copy = &H9C4C       ' 4 �R�s�[             9c4ch
                        ' 5 ------------
    Search = &H8006     ' 6 ����               8006h
                        ' 7 ------------
    PageOne = &H9C4E    ' 8 �P��y�[�W         9c4Eh
    PageMulti = &H9C4D  ' 9 �����y�[�W         9c4Dh
                        '10 ------------
    Reduce = &H9C46     '11 �k��               9c46h
    Expand = &H8007     '12 �g��               8007h
    ZoomRatio = &H8003  '13 �{��               8003h
                        '14 ------------
    BeforePage = &H9C44 '15 �O�y�[�W           9c44h
    NextPage = &H9C45   '16 ���y�[�W           9c45h
    GotoPage = &H8004   '17 �y�[�W�ԍ�         8004h
                        '18 ------------
    Back = &H8008       '19 �߂�               8008h
    Forward = &H9C4A    '20 �i��               9c4Ah
#End If
End Enum


Public Sub Setup(vReport As ActiveReport, Optional vPaperSize As Integer = vbPRPSA4, Optional vOrientation As Integer = vbPRORLandscape, Optional vImage As Variant = Empty)
    Set mReport = vReport
    With mReport
        '.WindowState = 2   '//���̐ݒ�͎��s���o���Ȃ��̂Ńf�U�C�����ɐݒ肵�Ă�������
        .Zoom = -2
        .Printer.PaperSize = vPaperSize
        .Printer.Orientation = vOrientation '// vbPRORLandscape = ������ : vbPRORPortrait = �c����
#If ARREPORT_VERSION = "1.5" Then
        .PageRightMargin = 10   '//�E�̓t���Ɉ��
        .PageBottomMargin = 10 '//���̓t���Ɉ��
#ElseIf ARREPORT_VERSION = "2.0" Then
        .PageSettings.RightMargin = 10   '//�E�̓t���Ɉ��
        .PageSettings.BottomMargin = 10  '//���̓t���Ɉ��
#End If
        '//Toolbar ���\��
        .Toolbar.Tools(eToolbarID.Contents).Visible = False    '���o��
'        .Toolbar.Tools(eToolbarID.Contents + 1).Visible = False  '���E
'        .Toolbar.Tools(eToolbarID.GotoPage + 1).Visible = False  '���E
        .Toolbar.Tools(eToolbarID.Back).Visible = False   '�߂�
        .Toolbar.Tools(eToolbarID.Forward).Visible = False   '�i��
        '//��ɂ���Ɣz�񂪂����
        Call .Toolbar.Tools.Insert(eToolbarID.Contents + 1, "����ݒ�(&S)")
'//�A�C�R����ݒ肵�悤�Ƃ��������܂������Ȃ��̂Ŏ~�߁I�I
'        If Not IsEmpty(vImage) Then
'            Call .Toolbar.Tools.Item(eToolbarID.Contents + 1).AddIcon(vImage)
'        End If
        .Toolbar.Tools(eToolbarID.Contents + 1).ID = eToolbar.Settings
'        With .Toolbar
'            Dim ix As Integer
'            For ix = 0 To .Tools.Count - 1
'                .Tools(ix).Visible = False
'            Next ix
'            Call .Tools.Add("���(&P)")
'            .Tools(.Tools.Count - 1).ID = 5113
'        End With
    End With
End Sub

Private Sub Class_Initialize()
    Set mReport = New ActiveReport
End Sub

Private Sub Class_Terminate()
    Set mReport = Nothing
End Sub

Private Sub mReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload mReport
    End If
End Sub

Private Sub mReport_NoData()
    Call MsgBox("�Y������f�[�^�͂���܂���.", vbInformation, mReport.documentName)
    Unload mReport
End Sub

#If ARREPORT_VERSION = "1.5" Then
Private Sub mReport_ToolbarClick(ByVal tool As DDActiveReports.DDTool)
#ElseIf ARREPORT_VERSION = "2.0" Then
Private Sub mReport_ToolbarClick(ByVal tool As DDActiveReports2.DDTool)
#End If
    Select Case tool.ID
    Case eToolbar.PrintOut
        Call mReport.Run
    Case eToolbar.Settings
#If 1 Then
        Call mReport.PageSetup              'ActiveReport PrintDialog
#Else
        '//���̕������Ƃ��܂��ĕ\�����Ȃ�
        Call mReport.Printer.PrintDialog    'VB PrintDialog
#End If
        Call mReport.Restart
    End Select
End Sub
