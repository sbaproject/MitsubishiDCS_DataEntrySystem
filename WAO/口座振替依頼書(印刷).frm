VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmKouzaFurikaeIraishoPrint 
   Caption         =   "�����U�ֈ˗���(���)"
   ClientHeight    =   4725
   ClientLeft      =   3750
   ClientTop       =   1800
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6615
   Begin VB.Frame fraSort 
      Caption         =   "�o�͏���"
      Height          =   945
      Left            =   1305
      TabIndex        =   14
      Top             =   2565
      Width           =   1965
      Begin VB.OptionButton optSort 
         Caption         =   "�f�[�^���� ��"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optSort 
         Caption         =   "�_��Ҕԍ� ��"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   270
         Width           =   1575
      End
   End
   Begin VB.Frame fraImport 
      BackColor       =   &H000000FF&
      Caption         =   "�Ώێ�(�捞��)"
      Height          =   1035
      Left            =   3420
      TabIndex        =   11
      Top             =   1380
      Visible         =   0   'False
      Width           =   1695
      Begin VB.CheckBox chkTaisho 
         Caption         =   "�V�K�o�^��"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkTaisho 
         Caption         =   "�C����"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame fraInput 
      Caption         =   "�Ώێ�(����͕�)"
      Height          =   1035
      Left            =   1320
      TabIndex        =   8
      Top             =   1380
      Width           =   1695
      Begin VB.CheckBox chkTaisho 
         Caption         =   "�C����"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   600
         Value           =   1  '����
         Width           =   1335
      End
      Begin VB.CheckBox chkTaisho 
         Caption         =   "�V�K�o�^��"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   240
         Value           =   1  '����
         Width           =   1455
      End
   End
   Begin MSDBCtls.DBCombo cboItakusha 
      Bindings        =   "�����U�ֈ˗���(���).frx":0000
      Height          =   300
      Left            =   1920
      TabIndex        =   6
      Top             =   900
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   529
      _Version        =   393216
      Style           =   2
      ListField       =   "ABKJNM"
      BoundColumn     =   "ABITKB"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkDefault 
      Caption         =   "�O��ݐϓ�"
      Height          =   315
      Left            =   3900
      TabIndex        =   5
      Top             =   420
      Width           =   1875
   End
   Begin imText6Ctl.imText txtStartDate 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   420
      Width           =   1875
      _Version        =   65536
      _ExtentX        =   3307
      _ExtentY        =   556
      Caption         =   "�����U�ֈ˗���(���).frx":002C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�����U�ֈ˗���(���).frx":009A
      Key             =   "�����U�ֈ˗���(���).frx":00B8
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   "2004/06/28 12:13:14"
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   3
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "���(&P)"
      Height          =   435
      Left            =   450
      TabIndex        =   1
      ToolTipText     =   "������J�n����ꍇ"
      Top             =   3915
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "�I��(&E)"
      Default         =   -1  'True
      Height          =   435
      Left            =   4710
      TabIndex        =   0
      ToolTipText     =   "���̍�Ƃ��I�����ă��C�����j���[�ɖ߂�ꍇ"
      Top             =   3915
      Width           =   1335
   End
   Begin ORADCLibCtl.ORADC dbcItakushaMaster 
      Height          =   315
      Left            =   2430
      Top             =   3975
      Visible         =   0   'False
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
      _ExtentY        =   556
      _StockProps     =   207
      Caption         =   "taItakushaMaster"
      BackColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.01
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   "dcssvr03"
      Connect         =   "kumon/kumon"
      RecordSource    =   "SELECT ABITKB,ABKJNM FROM taItakushaMaster"
   End
   Begin VB.Label Label2 
      Caption         =   "�ϑ���"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "���"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label1"
      Height          =   195
      Left            =   4860
      TabIndex        =   2
      Top             =   0
      Width           =   1275
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
Attribute VB_Name = "frmKouzaFurikaeIraishoPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm As New FormClass
Private mCaption As String
Private mStartDate As String
Private mYubinCode As String
Private mYubinName As String

Private Enum eSort
    eKeiyakusha = 0
    eInput
End Enum

Private Enum eTaisho
    eNewInput       '//�V�K�����
    eEditInput      '//�C�������
    eNewImport      '//�V�K�捞
    eEditImport     '//�C���捞
End Enum

Private Sub cboItakusha_Click(Area As Integer)
    Select Case Area
    Case 1
    Case dbcAreaButton      '// 0 DB �R���{ �R���g���[����Ń{�^�����N���b�N����܂����B
    Case dbcAreaEdit        '// 1 DB �R���{ �R���g���[���̃e�L�X�g �{�b�N�X���N���b�N����܂����B
    Case dbcAreaList        '// 2 DB �R���{ �R���g���[���̃h���b�v�_�E�� ���X�g �{�b�N�X���N���b�N����܂����B
        Debug.Print
    End Select
End Sub

Private Sub chkDefault_Click()
    If 0 = chkDefault.Value Then
        txtStartDate.Enabled = True
    Else
        txtStartDate.Text = mStartDate
        txtStartDate.Enabled = False
    End If
End Sub

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Function pCheckDate(vDate As Variant) As Variant
    On Error GoTo pCheckDateError:
    pCheckDate = CVDate(vDate)
    Exit Function
pCheckDateError:
    Call MsgBox("�w�肳�ꂽ������s���ł��B", vbCritical + vbOKOnly, mCaption)
End Function

Private Sub cmdPrint_Click()
    Dim StartDate As Variant
    '//Oracle �� Format �ɕϊ�����K�v������
    If "" <> Trim(txtStartDate.Text) Then
        StartDate = Format(pCheckDate(txtStartDate.Text), "YYYY/MM/DD HH:NN:SS")
        If Not IsDate(StartDate) Then
            Exit Sub
        End If
    End If
    If chkTaisho(eTaisho.eNewInput).Value = 0 And chkTaisho(eTaisho.eEditInput).Value = 0 _
    And chkTaisho(eTaisho.eNewImport).Value = 0 And chkTaisho(eTaisho.eEditImport).Value = 0 Then
        Call MsgBox("�Ώێ҂��I������Ă��܂���B", vbCritical + vbOKOnly, mCaption)
        Exit Sub
    End If
    Dim sql As String
    sql = "SELECT a.*," & vbCrLf
    sql = sql & " DECODE(CAKKBN,0,NULL,1,'�X','��') CAKKBNx," & vbCrLf
    sql = sql & " DECODE(CAKKBN,0,DECODE(CAKZSB,1,'��',2,'��','��'),NULL) CAKZSBx," & vbCrLf
    sql = sql & " DECODE(CAKYFG,0,NULL,1,'���','����') CAKYFGx," & vbCrLf
    sql = sql & " b.DAKJNM BankName," & vbCrLf
    sql = sql & " c.DAKJNM ShitenName," & vbCrLf
    sql = sql & " d.ABKJNM," & vbCrLf
    sql = sql & " a.CAUPDT INPDATE," & vbCrLf
    sql = sql & " a.CAUSID INPUSER " & vbCrLf
    sql = sql & " FROM tcHogoshaMaster  a," & vbCrLf
    sql = sql & "      tdBankMaster     b," & vbCrLf
    sql = sql & "      tdBankMaster     c," & vbCrLf
    sql = sql & "      taItakushaMaster d " & vbCrLf
    sql = sql & " WHERE CABANK = b.DABANK(+)" & vbCrLf
    sql = sql & "   AND '000'  = b.DASITN(+)" & vbCrLf
    sql = sql & "   AND ':'    = b.DASQNO(+)" & vbCrLf
    sql = sql & "   AND CABANK = c.DABANK(+)" & vbCrLf
    sql = sql & "   AND CASITN = c.DASITN(+)" & vbCrLf
    sql = sql & "   AND '�'    = c.DASQNO(+)" & vbCrLf
    sql = sql & "   AND CAITKB = ABITKB " & vbCrLf
    If -1 <> cboItakusha.BoundText Then
        sql = sql & "   AND CAITKB = " & cboItakusha.BoundText & vbCrLf
    End If
    If IsDate(StartDate) Then
        '///////////////////////////
        If 0 <> chkTaisho(eTaisho.eNewInput).Value And 0 <> chkTaisho(eTaisho.eEditInput).Value Then
            If 0 <> chkTaisho(eTaisho.eNewImport).Value And 0 <> chkTaisho(eTaisho.eEditImport).Value Then
                '//����̓f�[�^/�V�K/�ύX �� �捞�f�[�^/�V�K/�ύX�F�S��
                sql = sql & "   AND(CAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "    OR CAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "   )"
            ElseIf 0 <> chkTaisho(eTaisho.eNewImport).Value Then
                '//����̓f�[�^/�V�K/�ύX �� �捞�f�[�^/�V�K
                sql = sql & "   AND(CAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "    OR(" & vbCrLf
                sql = sql & "           CAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                '//�C���̎捞��(USER=PUNCH_IMPORT)�ȊO
                sql = sql & "       AND CAUSID <> " & gdDBS.ColumnDataSet(MainModule.gcImportUserName, vEnd:=True) & vbCrLf
                sql = sql & "      )"
                sql = sql & "   )"
            ElseIf 0 <> chkTaisho(eTaisho.eEditImport).Value Then
                '//����̓f�[�^/�V�K/�ύX �� �捞�f�[�^/�ύX
                sql = sql & "   AND((" & vbCrLf
                sql = sql & "           CAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                '//�V�K�̎捞��(USER=PUNCH_IMPORT)�ȊO
                sql = sql & "       AND CAUSID <> " & gdDBS.ColumnDataSet(MainModule.gcImportUserName, vEnd:=True) & vbCrLf
                sql = sql & "      )"
                sql = sql & "    OR CAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "   )"
            Else
                '//����̓f�[�^/�V�K/�ύX �� �捞�f�[�^/����
                sql = sql & "   AND(CAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "    OR CAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "   )"
                sql = sql & "   AND CAUSID <> " & gdDBS.ColumnDataSet(MainModule.gcImportUserName, vEnd:=True) & vbCrLf
            End If
        ElseIf 0 <> chkTaisho(eTaisho.eNewInput).Value Then
            If 0 <> chkTaisho(eTaisho.eNewImport).Value And 0 <> chkTaisho(eTaisho.eEditImport).Value Then
                '//����̓f�[�^/�V�K �� �捞�f�[�^/�V�K/�ύX
                sql = sql & "   AND CAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "    OR( CAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "    AND CAUSID = " & gdDBS.ColumnDataSet(MainModule.gcImportUserName, vEnd:=True) & vbCrLf
                sql = sql & "   )"
            ElseIf 0 <> chkTaisho(eTaisho.eNewImport).Value Then
                '//����̓f�[�^/�V�K �� �捞�f�[�^/�V�K
                sql = sql & "   AND CAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
            ElseIf 0 <> chkTaisho(eTaisho.eEditImport).Value Then
                '//����̓f�[�^/�V�K �� �捞�f�[�^/�ύX
                sql = sql & "   AND(CAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "   AND CAUSID <> " & gdDBS.ColumnDataSet(MainModule.gcImportUserName, vEnd:=True) & vbCrLf
                sql = sql & "   )"
                sql = sql & "   AND(CAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "   AND CAUSID =  " & gdDBS.ColumnDataSet(MainModule.gcImportUserName, vEnd:=True) & vbCrLf
                sql = sql & "   )"
            Else
                '//����̓f�[�^/�V�K �� �捞�f�[�^/����
                sql = sql & "   AND CAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "   AND CAUSID <> " & gdDBS.ColumnDataSet(MainModule.gcImportUserName, vEnd:=True) & vbCrLf
            End If
        ElseIf 0 <> chkTaisho(eTaisho.eEditInput).Value Then
            If 0 <> chkTaisho(eTaisho.eNewImport).Value And 0 <> chkTaisho(eTaisho.eEditImport).Value Then
                '//����̓f�[�^/�C�� �� �捞�f�[�^/�V�K/�ύX
                sql = sql & "   AND CAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "    OR( CAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "    AND CAUSID = " & gdDBS.ColumnDataSet(MainModule.gcImportUserName, vEnd:=True) & vbCrLf
                sql = sql & "   )"
            ElseIf 0 <> chkTaisho(eTaisho.eNewImport).Value Then
                '//����̓f�[�^/�C�� �� �捞�f�[�^/�V�K
                sql = sql & "   AND(" & vbCrLf
                sql = sql & "         ( CAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "       AND CAUSID <> " & gdDBS.ColumnDataSet(MainModule.gcImportUserName, vEnd:=True) & vbCrLf
                sql = sql & "      )OR( CAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "       AND CAUSID =  " & gdDBS.ColumnDataSet(MainModule.gcImportUserName, vEnd:=True) & vbCrLf
                sql = sql & "      )" & vbCrLf
                sql = sql & "   )" & vbCrLf
            ElseIf 0 <> chkTaisho(eTaisho.eEditImport).Value Then
                '//����̓f�[�^/�C�� �� �捞�f�[�^/�ύX
                sql = sql & "   AND CAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
            Else
                '//����̓f�[�^/�C�� �� �捞�f�[�^/����
                sql = sql & "   AND CAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
                sql = sql & "   AND CAUSID <> " & gdDBS.ColumnDataSet(MainModule.gcImportUserName, vEnd:=True) & vbCrLf
            End If
        ElseIf 0 <> chkTaisho(eTaisho.eNewImport).Value And 0 <> chkTaisho(eTaisho.eEditImport).Value Then
            '//�捞�f�[�^/�V�K/�ύX
            sql = sql & "   AND(CAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
            sql = sql & "    OR CAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
            sql = sql & "   )"
            sql = sql & "   AND CAUSID = " & gdDBS.ColumnDataSet(MainModule.gcImportUserName, vEnd:=True) & vbCrLf
        ElseIf 0 <> chkTaisho(eTaisho.eNewImport).Value Then
            '//�捞�f�[�^/�V�K
            sql = sql & "   AND CAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
            sql = sql & "   AND CAUSID = " & gdDBS.ColumnDataSet(MainModule.gcImportUserName, vEnd:=True) & vbCrLf
        ElseIf 0 <> chkTaisho(eTaisho.eEditImport).Value Then
            '//�捞�f�[�^/�ύX
            sql = sql & "   AND CAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
            sql = sql & "   AND CAUSID = " & gdDBS.ColumnDataSet(MainModule.gcImportUserName, vEnd:=True) & vbCrLf
        End If
    End If      '// If IsDate(StartDate) Then
    'sql = sql & " ORDER BY CAITKB,CAKYCD,CAHGCD,CASQNO"
    Select Case Val(fraSort.Tag)
    Case eSort.eKeiyakusha
        sql = sql & " ORDER BY CAITKB,CAKYCD,CAHGCD,CASQNO" & vbCrLf
    Case eSort.eInput
        sql = sql & " ORDER BY INPDATE,CAITKB,CAKYCD,CAHGCD,CASQNO" & vbCrLf
    End Select
    Dim reg As New RegistryClass
    Load rptKouzaFurikaeIraisho
    With rptKouzaFurikaeIraisho
        .lblCondition.Caption = ""
        If 0 <> chkDefault.Value Then
            .lblCondition.Caption = "����F" & chkDefault.Caption
        ElseIf "" <> Trim(txtStartDate.Text) Then
            .lblCondition.Caption = "����F" & txtStartDate.Text
        End If
        .lblCondition.Caption = .lblCondition.Caption & "  �o�͏��ԁF" & optSort(Val(fraSort.Tag)).Caption
        .lblCondition.Caption = .lblCondition.Caption & "  "
        If 0 <> chkTaisho(eTaisho.eNewInput).Value And 0 <> chkTaisho(eTaisho.eEditInput).Value Then
            .lblCondition.Caption = .lblCondition.Caption & fraInput.Caption & "�F" & chkTaisho(eTaisho.eNewInput).Caption & "��" & chkTaisho(eTaisho.eEditInput).Caption
        ElseIf 0 <> chkTaisho(eTaisho.eNewInput).Value Then
            .lblCondition.Caption = .lblCondition.Caption & fraInput.Caption & "�F" & chkTaisho(eTaisho.eNewInput).Caption
        ElseIf 0 <> chkTaisho(eTaisho.eEditInput).Value Then
            .lblCondition.Caption = .lblCondition.Caption & fraInput.Caption & "�F" & chkTaisho(eTaisho.eEditInput).Caption
        End If
        .lblCondition.Caption = .lblCondition.Caption & "  "
        If 0 <> chkTaisho(eTaisho.eNewImport).Value And 0 <> chkTaisho(eTaisho.eEditImport).Value Then
            .lblCondition.Caption = .lblCondition.Caption & fraImport.Caption & "�F" & chkTaisho(eTaisho.eNewImport).Caption & "��" & chkTaisho(eTaisho.eEditImport).Caption
        ElseIf 0 <> chkTaisho(eTaisho.eNewImport).Value Then
            .lblCondition.Caption = .lblCondition.Caption & fraImport.Caption & "�F" & chkTaisho(eTaisho.eNewImport).Caption
        ElseIf 0 <> chkTaisho(eTaisho.eEditImport).Value Then
            .lblCondition.Caption = .lblCondition.Caption & fraImport.Caption & "�F" & chkTaisho(eTaisho.eEditImport).Caption
        End If
        .mStartDate = mStartDate
        .mYubinCode = mYubinCode
        .mYubinName = mYubinName
        .documentName = mCaption
        .adoData.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=" & reg.DbPassword & _
                                    ";Persist Security Info=True;User ID=" & reg.DbUserName & _
                                                           ";Data Source=" & reg.DbDatabaseName
        .adoData.Source = sql
        'Call .adoData.Refresh
        Call .Show
    End With
End Sub

Private Sub Form_Activate()
    If "" = Trim(cboItakusha.BoundText) Then
        cboItakusha.BoundText = "-1"
    End If
End Sub

Private Sub optSort_Click(Index As Integer)
    fraSort.Tag = Index
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    Call mForm.LockedControl(False)
    fraImport.Visible = False           '//�������͎捞�ݖ���
    Dim sql As String, dyn As Object
    sql = "SELECT * FROM taSystemInformation"
    Set dyn = gdDBS.OpenRecordset(sql)
    If dyn.EOF Then
        mStartDate = Now()
    Else
        mStartDate = Format(dyn.Fields("AANWDT").Value, "yyyy/mm/dd hh:nn:ss")
        mYubinCode = dyn.Fields("AAYSNO").Value
        mYubinName = dyn.Fields("AAYSNM").Value
    End If
    Call dyn.Close
    txtStartDate.Text = mStartDate
    
    optSort(0).Value = True
    
    sql = "SELECT * FROM("
    sql = sql & "SELECT '-1' ABITKB,'<< �S�Ă�Ώ� >>' ABKJNM FROM DUAL"
    sql = sql & " UNION "
    sql = sql & "SELECT ABITKB,ABKJNM FROM taItakushaMaster"
    sql = sql & ")"
    dbcItakushaMaster.RecordSource = sql
    Call dbcItakushaMaster.Refresh
    chkDefault.Value = 1
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmKouzaFurikaeIraishoPrint = Nothing
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

