VERSION 5.00
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmKeiyakushaMasterExport 
   Caption         =   "�_��҃}�X�^�f�[�^�쐬"
   ClientHeight    =   3825
   ClientLeft      =   2865
   ClientTop       =   4035
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6180
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   540
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin imDate6Ctl.imDate txtBAKYED 
      Height          =   285
      Left            =   2340
      TabIndex        =   0
      Top             =   120
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   503
      Calendar        =   "�_��҃}�X�^�f�[�^�쐬.frx":0000
      Caption         =   "�_��҃}�X�^�f�[�^�쐬.frx":0186
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�f�[�^�쐬.frx":01F4
      Keys            =   "�_��҃}�X�^�f�[�^�쐬.frx":0212
      MouseIcon       =   "�_��҃}�X�^�f�[�^�쐬.frx":0270
      Spin            =   "�_��҃}�X�^�f�[�^�쐬.frx":028C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "yyyy/mm/dd"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy/mm/dd"
      HighlightText   =   2
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   73050
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   " "
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "    /  /  "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   -2
      CenturyMode     =   0
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Text�쐬(�z�X�g����)(&E)"
      Height          =   555
      Left            =   600
      TabIndex        =   4
      Top             =   3060
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport2 
      Caption         =   "CSV�쐬(�c�e����) (&S)"
      Height          =   555
      Left            =   2340
      TabIndex        =   5
      Top             =   3060
      Width           =   1455
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "�I��(&X)"
      Height          =   555
      Left            =   4020
      TabIndex        =   6
      Top             =   3060
      Width           =   1455
   End
   Begin imDate6Ctl.imDate txtNewData 
      Height          =   285
      Left            =   2340
      TabIndex        =   1
      Top             =   540
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   503
      Calendar        =   "�_��҃}�X�^�f�[�^�쐬.frx":02B4
      Caption         =   "�_��҃}�X�^�f�[�^�쐬.frx":043A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^�f�[�^�쐬.frx":04A8
      Keys            =   "�_��҃}�X�^�f�[�^�쐬.frx":04C6
      MouseIcon       =   "�_��҃}�X�^�f�[�^�쐬.frx":0524
      Spin            =   "�_��҃}�X�^�f�[�^�쐬.frx":0540
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "yyyy/mm/dd"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy/mm/dd"
      HighlightText   =   2
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   73050
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   " "
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "    /  /  "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   -2
      CenturyMode     =   0
   End
   Begin VB.Label lblMessageB 
      Caption         =   "�c�e�n ���b�Z�[�W"
      Height          =   915
      Left            =   780
      TabIndex        =   11
      Top             =   2040
      Width           =   5355
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�y�쐬�菇�z"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   900
      TabIndex        =   10
      Top             =   1020
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "�ȍ~��V�K�����Ƃ���B"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "�V�K���"
      Height          =   255
      Left            =   1380
      TabIndex        =   9
      Top             =   600
      Width           =   915
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label1"
      Height          =   195
      Left            =   4140
      TabIndex        =   8
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label Label8 
      Caption         =   "�_��L����"
      Height          =   255
      Left            =   1380
      TabIndex        =   7
      Top             =   180
      Width           =   915
   End
   Begin VB.Label lblMessageA 
      Caption         =   "�z�X�g�n ���b�Z�[�W"
      Height          =   675
      Left            =   780
      TabIndex        =   3
      Top             =   1320
      Width           =   5355
   End
   Begin VB.Menu mnuFile 
      Caption         =   "̧��(&F)"
      Begin VB.Menu mnuEnd 
         Caption         =   "�I��(&X)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuVersion 
         Caption         =   "�ް�ޮݏ��(&A)"
      End
   End
End
Attribute VB_Name = "frmKeiyakushaMasterExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCaption As String
Private Const mExeMsgA As String = "�z�X�g���� Text�쐬" & vbCrLf & _
                                "�@�z�X�g�ւ̌Œ蒷�e�L�X�g�쐬�����܂��B" & vbCrLf
Private Const mExeMsgB As String = "�c�e���� CSV�쐬" & vbCrLf & _
                                "�@�c�e�ւ̂b�r�u�e�L�X�g�쐬�����܂��B" & vbCrLf & _
                                "�@�����F�_��L�����A�V�K����͖�������܂��B" & vbCrLf
Private mForm As New FormClass
Private mAbort As Boolean

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    '//���b�Z�[�W������
    lblMessageA.Caption = mExeMsgA
    lblMessageA.FontBold = False
    lblMessageA.ForeColor = vbBlack
    lblMessageB.Caption = mExeMsgB
    lblMessageB.FontBold = False
    lblMessageB.ForeColor = vbBlack
    Const cFileID As String = "�z�X�g����."
'    On Error GoTo cmdExport_ClickError
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset, dyn2 As OraDynaset
#Else
    Dim sql As String, dyn As Object, dyn2 As Object
#End If
    
    sql = "SELECT * "
    sql = sql & " FROM taItakushaMaster,"
    sql = sql & "      tbKeiyakushaMaster"
    sql = sql & " WHERE ABITKB = BAITKB"
    '//�_������L���͈͂��H
    sql = sql & "   AND " & txtBAKYED.Number & " BETWEEN BAKYST AND BAKYED"
    '//�U�֓��̗L���͈͂��H
    sql = sql & "   AND " & txtBAKYED.Number & " BETWEEN BAFKST AND BAFKED"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If dyn.EOF Then
        Call MsgBox("�Y������f�[�^�͂���܂���.", vbInformation + vbOKOnly, cFileID & mCaption)
        Exit Sub
    End If
    Dim st As New StructureClass, tmp As String
    Dim reg As New RegistryClass
    Dim mFile As New FileClass, FileName As String, TmpFname As String
    Dim fp As Integer, cnt As Long
    
    dlgFile.DialogTitle = "���O��t���ĕۑ�(" & cFileID & mCaption & ")"
    dlgFile.FileName = reg.OutputFileName(cFileID & mCaption)
    If IsEmpty(mFile.SaveDialog(dlgFile)) Then
        Exit Sub
    End If
    
    Dim ms As New MouseClass
    Call ms.Start
    
    reg.OutputFileName(cFileID & mCaption) = dlgFile.FileName
    Call st.SelectStructure(st.Keiyakusha)
    
    '//��芸�����e���|�����ɏ���
    fp = FreeFile
    TmpFname = mFile.MakeTempFile
    Open TmpFname For Append As #fp
    Do Until dyn.EOF
        DoEvents
        If mAbort Then
            GoTo cmdExport_ClickAbort
        End If
        tmp = ""
        tmp = tmp & st.SetData(dyn.Fields("ABITCD"), 0)     '�ϑ��Ҕԍ�             '//���̍��ڂ͈ϑ��҃}�X�^
        tmp = tmp & st.SetData(dyn.Fields("BAKYCD"), 1)     '�_��Ҕԍ�(����)
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//        tmp = tmp & st.SetData(dyn.Fields("BAKSCD"), 2)     '�����敪
        tmp = tmp & st.SetData("000", 2)     '�����敪
        '//2002/11/26 �󔒂T�����ǉ�
        tmp = tmp & String(5, " ")
        '//���Z�@�ւ̋敪�ɂ���ċ�s���X�֋ǂ̌��ʂ�ԋp����֐��� StructureClass ���쐬
        tmp = tmp & st.SetData(st.BankCode(dyn), 3)         '��s�R�[�h
        tmp = tmp & st.SetData(st.ShitenCode(dyn), 4)       '�x�X�R�[�h
        tmp = tmp & st.SetData(st.Shubetsu(dyn), 5)         '�a�����
        tmp = tmp & st.SetData(st.KouzaNo(dyn), 6)          '�����ԍ�
        '//���Z�@�ւ̋敪�ɂ���ċ�s���X�֋ǂ̌��ʂ�ԋp����֐��� StructureClass ���쐬
        tmp = tmp & st.SetData(dyn.Fields("BAKZNM"), 7)     '�������`�l��(�J�i)
        tmp = tmp & st.SetData(dyn.Fields("BAZPC1"), 8)     '�X�֔ԍ��P
        tmp = tmp & st.SetData(dyn.Fields("BAZPC2"), 9)     '�X�֔ԍ��Q
        tmp = tmp & st.SetData(dyn.Fields("BAADJ1"), 10)    '�Z���P(����)
        tmp = tmp & st.SetData(dyn.Fields("BAADJ2"), 11)    '�Z���Q(����)
        tmp = tmp & st.SetData(dyn.Fields("BAADJ3"), 12)    '�Z���R(����)
        tmp = tmp & st.SetData(dyn.Fields("BAKJNM"), 13)    '����
        tmp = tmp & st.SetData(dyn.Fields("BAKSNO"), 14)    '�����ԍ�
        tmp = tmp & st.SetData(dyn.Fields("BATELE"), 15)    '�d�b�ԍ��P     (����)
        tmp = tmp & st.SetData(dyn.Fields("BATELJ"), 16)    '�d�b�ԍ��Q     (����)
        tmp = tmp & st.SetData(dyn.Fields("BAKKRN"), 17)    '�d�b�ԍ��R     (�ً})
        tmp = tmp & st.SetData(dyn.Fields("BAFAXI"), 18)    '�e�`�w�ԍ��P   (����)
        tmp = tmp & st.SetData(dyn.Fields("BAFAXJ"), 19)    '�e�`�w�ԍ��Q   (����)
        '//�ی�҂̐V�K�l�����J�E���g����
        sql = "SELECT COUNT(*) AS CNT"
        sql = sql & " FROM tcHogoshaMaster"
        sql = sql & " WHERE CAITKB = '" & CStr(dyn.Fields("BAITKB")) & "'"
        sql = sql & "   AND CAKYCD = '" & CStr(dyn.Fields("BAKYCD")) & "'"
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//        sql = sql & "   AND CAKSCD = '" & dyn.Fields("BAKSCD") & "'"
        '//�_��J�n�I�����͈͓����H
        sql = sql & "   AND " & txtBAKYED.Number & " BETWEEN CAKYST AND CAKYED"
        '//�_��J�n�����L�����ȉ����H
        sql = sql & "   AND CAADDT >= TO_DATE('" & txtNewData.Text & " 00:00:00','YYYY/MM/DD HH24:MI:SS')"
#If ORA_DEBUG = 1 Then
        Set dyn2 = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
        Set dyn2 = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
        tmp = tmp & st.SetData(dyn2.Fields("CNT"), 20)      '�ی�҂̐V�K�l��
'//2003/02/03 (21)�̍��ڂ��s�����Ă����̂Œǉ�
        tmp = tmp & st.SetData(0, 21)                       '�����z�F��������́H�H�H
        Call dyn2.Close
        Print #fp, tmp
        cnt = cnt + 1
        Call dyn.MoveNext
    Loop
    Call dyn.Close
#If NO_TOTAL_REC Then
'//2003/02/03 ��������@���v�����E���z���R�[�h�ǉ�
    tmp = ""
    tmp = tmp & st.SetData("9999999999", 0)     '�ϑ��Ҕԍ�             '//���̍��ڂ͈ϑ��҃}�X�^
    tmp = tmp & st.SetData("9999999999", 1)     '�_��Ҕԍ�(����)
    tmp = tmp & st.SetData("9999999999", 2)     '�����敪
    tmp = tmp & String(5, " ")                  '//2002/11/26 �󔒂T�����ǉ�
    tmp = tmp & st.SetData("", 3)               '��s�R�[�h
    tmp = tmp & st.SetData("", 4)               '�x�X�R�[�h
    tmp = tmp & st.SetData("", 5)               '�a�����
    tmp = tmp & st.SetData("", 6)               '�����ԍ�
    tmp = tmp & st.SetData("", 7)               '�������`�l��(�J�i)
    tmp = tmp & st.SetData("", 8)               '�X�֔ԍ��P
    tmp = tmp & st.SetData("", 9)               '�X�֔ԍ��Q
    tmp = tmp & st.SetData("���v���R�[�h", 10)  '�Z���P(����)
    tmp = tmp & st.SetData("", 11)              '�Z���Q(����)
    tmp = tmp & st.SetData("", 12)              '�Z���R(����)
    tmp = tmp & st.SetData("", 13)              '����
    tmp = tmp & st.SetData("", 14)              '�����ԍ�
    tmp = tmp & st.SetData("", 15)              '�d�b�ԍ��P     (����)
    tmp = tmp & st.SetData("", 16)              '�d�b�ԍ��Q     (����)
    tmp = tmp & st.SetData("", 17)              '�d�b�ԍ��R     (�ً})
    tmp = tmp & st.SetData("", 18)              '�e�`�w�ԍ��P   (����)
    tmp = tmp & st.SetData("", 19)              '�e�`�w�ԍ��Q   (����)
    tmp = tmp & st.SetData("", 20)              '�ی�҂̐V�K�l��
    tmp = tmp & st.SetData(cnt, 21)             '������ ���v���z ������ �ی�҂̐V�K�l��
    Print #fp, tmp
'//2003/02/03 �����܂Ł@���v�����E���z���R�[�h�ǉ�
#End If
    Close #fp
#If 1 Then
    '//�t�@�C���ړ�     MOVEFILE_REPLACE_EXISTING=Replace , MOVEFILE_COPY_ALLOWED=Copy & Delete
    Call MoveFileEx(TmpFname, reg.OutputFileName(cFileID & mCaption), MOVEFILE_REPLACE_EXISTING + MOVEFILE_COPY_ALLOWED)
    'Call MoveFileEx(TmpFname, reg.FileName(cFileID & mCaption), MOVEFILE_REPLACE_EXISTING)
#Else
    '//�t�@�C���R�s�[
    Call FileCopy(TmpFname, reg.FileName(cFileID & mCaption))
#End If
    Set mFile = Nothing
    lblMessageA.Caption = mExeMsgA & cnt & " ���̃f�[�^���쐬����܂����B"
    lblMessageA.FontBold = True
    lblMessageA.ForeColor = vbBlue
    Exit Sub
cmdExport_ClickAbort:
    lblMessageA.Caption = mExeMsgA & "���~����܂����B"
    lblMessageA.FontBold = True
    lblMessageA.ForeColor = vbRed
    Exit Sub
cmdExport_ClickError:
    Call gdDBS.ErrorCheck       '//�G���[�g���b�v
    Set mFile = Nothing
End Sub

Private Sub cmdExport2_Click()
    '//���b�Z�[�W������
    lblMessageA.Caption = mExeMsgA
    lblMessageA.FontBold = False
    lblMessageA.ForeColor = vbBlack
    lblMessageB.Caption = mExeMsgB
    lblMessageB.FontBold = False
    lblMessageB.ForeColor = vbBlack
    Const cFileID As String = "�c�e����."
    'On Error GoTo cmdExport2_ClickError
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset, dyn2 As OraDynaset
#Else
    Dim sql As String, dyn As Object, dyn2 As Object
#End If
    
    sql = "SELECT * "
    sql = sql & " FROM taItakushaMaster,"
    sql = sql & "      tbKeiyakushaMaster"
    sql = sql & " WHERE ABITKB = BAITKB"
    sql = sql & "   AND (BAITKB,BAKYCD,BASQNO) IN("
    sql = sql & "       SELECT BAITKB,BAKYCD,max(BASQNO)"
    sql = sql & "       FROM tbKeiyakushaMaster"
    sql = sql & "       GROUP BY BAITKB,BAKYCD"
    sql = sql & "   )"
    sql = sql & "   AND BAKJNM IS NOT NULL"     '//�_��Ҏ����� NULL ������̂Ŕr��
    sql = sql & "   AND BAKNNM IS NOT NULL"     '//�_��Ҏ����� NULL ������̂Ŕr��
'    '//�_������L���͈͂��H
'    sql = sql & "   AND " & txtBAKYED.Number & " BETWEEN BAKYST AND BAKYED"
'    '//�U�֓��̗L���͈͂��H
'    sql = sql & "   AND " & txtBAKYED.Number & " BETWEEN BAFKST AND BAFKED"
    sql = sql & " ORDER BY BAITKB,BAKYCD"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If dyn.EOF Then
        Call MsgBox("�Y������f�[�^�͂���܂���.", vbInformation + vbOKOnly, cFileID & mCaption)
        Exit Sub
    End If
    
    Dim reg As New RegistryClass
    Dim mFile As New FileClass, FileName As String, TmpFname As String
    Dim fp As Integer, cnt As Long
    
    dlgFile.DialogTitle = "���O��t���ĕۑ�(" & cFileID & mCaption & ")"
    dlgFile.FileName = reg.OutputFileName(cFileID & mCaption)
    If IsEmpty(mFile.SaveDialog(dlgFile)) Then
        Exit Sub
    End If
    
    Dim ms As New MouseClass
    Call ms.Start
    
    reg.OutputFileName(cFileID & mCaption) = dlgFile.FileName
    '//��芸�����e���|�����ɏ���
    fp = FreeFile
    TmpFname = mFile.MakeTempFile
    Open TmpFname For Append As #fp
    Do Until dyn.EOF
        DoEvents
        If mAbort Then
            GoTo cmdExport2_ClickAbort
        End If
        '//�Z���͂P�A�Q�A�R�̍��ڂ��������ĂP�Q�O�o�C�g�ŕԋp
        Write #fp, dyn.Fields("BAKYCD"), _
                   dyn.Fields("BAKJNM"), _
                   dyn.Fields("BAKNNM"), _
                   pJoinStrings(dyn.Fields("BAADJ1") & dyn.Fields("BAADJ2") & dyn.Fields("BAADJ3"), 120)
        cnt = cnt + 1
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    Close #fp
#If 1 Then
    '//�t�@�C���ړ�     MOVEFILE_REPLACE_EXISTING=Replace , MOVEFILE_COPY_ALLOWED=Copy & Delete
    Call MoveFileEx(TmpFname, reg.OutputFileName(cFileID & mCaption), MOVEFILE_REPLACE_EXISTING + MOVEFILE_COPY_ALLOWED)
    'Call MoveFileEx(TmpFname, reg.FileName(cFileID & mCaption), MOVEFILE_REPLACE_EXISTING)
#Else
    '//�t�@�C���R�s�[
    Call FileCopy(TmpFname, reg.FileName(cFileID & mCaption))
#End If
    Set mFile = Nothing
    lblMessageB.Caption = mExeMsgB & cnt & " ���̃f�[�^���쐬����܂����B"
    lblMessageB.FontBold = True
    lblMessageB.ForeColor = vbMagenta

    Exit Sub
cmdExport2_ClickAbort:
    lblMessageB.Caption = mExeMsgA & "���~����܂����B"
    lblMessageB.FontBold = True
    lblMessageB.ForeColor = vbRed
    Exit Sub
cmdExport2_ClickError:
    Call gdDBS.ErrorCheck       '//�G���[�g���b�v
    Set mFile = Nothing
End Sub

'// Variant �Ŏ󂯂Ȃ��� DBNull �ŃG���[�ƂȂ�
Private Function pJoinStrings(vString As Variant, vBytes As Integer) As String
    If 0 <> Len(vString) Then
        '//�S�Ă̔��p�S�p�X�y�[�X���폜�����
        'pJoinStrings = Trim(StrConv(LeftB(StrConv(Replace(Replace(vString, "�@", ""), " ", ""), vbFromUnicode), vBytes), vbUnicode))
        '//�O��̔��p�S�p�X�y�[�X���폜�����
        pJoinStrings = Trim(StrConv(LeftB(StrConv(vString, vbFromUnicode), vBytes), vbUnicode))
    End If
End Function


'Private Sub cmdSend_Click()
'    Dim reg As New RegistryClass
'    Call Shell(reg.TransferCommand(mCaption), vbNormalFocus)
'End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    cmdExport.Caption = "�z�X�g����" & vbCrLf & "Text�쐬(&H)"
    cmdExport2.Caption = "�c�e����" & vbCrLf & "CSV�쐬(&D)"
    lblMessageA.Caption = mExeMsgA
    lblMessageB.Caption = mExeMsgB
    txtBAKYED.Number = gdDBS.sysDate("YYYYMMDD")
    txtNewData.Number = gdDBS.sysDate("YYYYMMDD")
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mAbort = True
    Set frmKeiyakushaMasterExport = Nothing
    Set mForm = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub mnuEnd_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

Private Sub txtBAKYED_DropOpen(NoDefault As Boolean)
    txtBAKYED.Calendar.Holidays = gdDBS.Holiday(txtBAKYED.Year)
End Sub

Private Sub txtNewData_DropOpen(NoDefault As Boolean)
    txtNewData.Calendar.Holidays = gdDBS.Holiday(txtNewData.Year)
End Sub

