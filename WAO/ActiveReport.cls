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
    Settings = 1     '印刷設定

#If ARREPORT_VERSION = "1.5" Then
    Contents = 0        '0   5110    目次
                        '1   500 -
    PrintOut = 2        '2   5113    印刷
                        '3   500 -
    Reduce = 4          '4   5108    縮小
    Expand = 5          '5   5109    拡大
    ZoomRatio = 6       '6   5114    倍率
                        '7   500 -
    BeforePage = 8      '8   5106    前ページ
    NextPage = 9        '9   5107    次ページ
    GotoPage = 10       '10  5115    ページ
                        '11  500 -
    Back = 12           '12  5111    戻る
    Forward = 13        '13  5112    進む
#ElseIf ARREPORT_VERSION = "2.0" Then
                        '以下は､ V2.0 定義済みのツールID一覧です｡
    Contents = 0        ' 0 TOC（見出し一覧）  9c48h
                        ' 1 ------------
    PrintOut = 2        ' 2 印刷               8005h
                        ' 3 ------------
    Copy = 4            ' 4 コピー             9c4ch
                        ' 5 ------------
    Search = 6          ' 6 検索               8006h
                        ' 7 ------------
    PageOne = 8         ' 8 単一ページ         9c4Eh
    PageMulti = 9       ' 9 複数ページ         9c4Dh
                        '10 ------------
    Reduce = 11         '11 縮小               9c46h
    Expand = 12         '12 拡大               8007h
    ZoomRatio = 13      '13 倍率               8003h
                        '14 ------------
    BeforePage = 15     '15 前ページ           9c44h
    NextPage = 16       '16 次ページ           9c45h
    GotoPage = 17       '17 ページ番号         8004h
                        '18 ------------
    Back = 19           '19 戻る               8008h
    Forward = 20        '20 進む               9c4Ah
#End If
End Enum

Private Enum eToolbar
    Settings = 1001     '印刷設定

#If ARREPORT_VERSION = "1.5" Then
    Contents = 5110     '0   5110    目次
                        '1   500 -
    PrintOut = 5113     '2   5113    印刷
                        '3   500 -
    Reduce = 5108       '4   5108    縮小
    Expand = 5109       '5   5109    拡大
    ZoomRatio = 5114    '6   5114    倍率
                        '7   500 -
    BeforePage = 5106   '8   5106    前ページ
    NextPage = 5107     '9   5107    次ページ
    GotoPage = 5115     '10  5115    ページ
                        '11  500 -
    Back = 5111         '12  5111    戻る
    Forward = 5112      '13  5112    進む
#ElseIf ARREPORT_VERSION = "2.0" Then
                        '以下は､ V2.0 定義済みのツールID一覧です｡
    Contents = &H9C48   ' 0 TOC（見出し一覧）  9c48h
                        ' 1 ------------
    PrintOut = &H8005   ' 2 印刷               8005h
                        ' 3 ------------
    Copy = &H9C4C       ' 4 コピー             9c4ch
                        ' 5 ------------
    Search = &H8006     ' 6 検索               8006h
                        ' 7 ------------
    PageOne = &H9C4E    ' 8 単一ページ         9c4Eh
    PageMulti = &H9C4D  ' 9 複数ページ         9c4Dh
                        '10 ------------
    Reduce = &H9C46     '11 縮小               9c46h
    Expand = &H8007     '12 拡大               8007h
    ZoomRatio = &H8003  '13 倍率               8003h
                        '14 ------------
    BeforePage = &H9C44 '15 前ページ           9c44h
    NextPage = &H9C45   '16 次ページ           9c45h
    GotoPage = &H8004   '17 ページ番号         8004h
                        '18 ------------
    Back = &H8008       '19 戻る               8008h
    Forward = &H9C4A    '20 進む               9c4Ah
#End If
End Enum


Public Sub Setup(vReport As ActiveReport, Optional vPaperSize As Integer = vbPRPSA4, Optional vOrientation As Integer = vbPRORLandscape, Optional vImage As Variant = Empty)
    Set mReport = vReport
    With mReport
        '.WindowState = 2   '//この設定は実行時出来ないのでデザイン時に設定しておくこと
        .Zoom = -2
        .Printer.PaperSize = vPaperSize
        .Printer.Orientation = vOrientation '// vbPRORLandscape = 横方向 : vbPRORPortrait = 縦方向
#If ARREPORT_VERSION = "1.5" Then
        .PageRightMargin = 10   '//右はフルに印刷
        .PageBottomMargin = 10 '//下はフルに印刷
#ElseIf ARREPORT_VERSION = "2.0" Then
        .PageSettings.RightMargin = 10   '//右はフルに印刷
        .PageSettings.BottomMargin = 10  '//下はフルに印刷
#End If
        '//Toolbar を非表示
        .Toolbar.Tools(eToolbarID.Contents).Visible = False    '見出し
'        .Toolbar.Tools(eToolbarID.Contents + 1).Visible = False  '境界
'        .Toolbar.Tools(eToolbarID.GotoPage + 1).Visible = False  '境界
        .Toolbar.Tools(eToolbarID.Back).Visible = False   '戻る
        .Toolbar.Tools(eToolbarID.Forward).Visible = False   '進む
        '//先にすると配列がずれる
        Call .Toolbar.Tools.Insert(eToolbarID.Contents + 1, "印刷設定(&S)")
'//アイコンを設定しようとしたがうまくいかないので止め！！
'        If Not IsEmpty(vImage) Then
'            Call .Toolbar.Tools.Item(eToolbarID.Contents + 1).AddIcon(vImage)
'        End If
        .Toolbar.Tools(eToolbarID.Contents + 1).ID = eToolbar.Settings
'        With .Toolbar
'            Dim ix As Integer
'            For ix = 0 To .Tools.Count - 1
'                .Tools(ix).Visible = False
'            Next ix
'            Call .Tools.Add("印刷(&P)")
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
    Call MsgBox("該当するデータはありません.", vbInformation, mReport.documentName)
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
        '//この方式だとうまく再表示しない
        Call mReport.Printer.PrintDialog    'VB PrintDialog
#End If
        Call mReport.Restart
    End Select
End Sub

