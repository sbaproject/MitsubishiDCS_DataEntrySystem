VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmKeiyakushaCheckList 
   Caption         =   "�I�[�i�[�}�X�^�`�F�b�N���X�g"
   ClientHeight    =   4050
   ClientLeft      =   3795
   ClientTop       =   2370
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6360
   Begin VB.ComboBox cboSort 
      Height          =   300
      ItemData        =   "�_��҃}�X�^(���).frx":0000
      Left            =   1980
      List            =   "�_��҃}�X�^(���).frx":000D
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   12
      Top             =   2580
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "�Ώێ�"
      Height          =   975
      Left            =   1920
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
      Bindings        =   "�_��҃}�X�^(���).frx":0033
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
      Width           =   1575
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
      Caption         =   "�_��҃}�X�^(���).frx":005F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�_��҃}�X�^(���).frx":00CD
      Key             =   "�_��҃}�X�^(���).frx":00EB
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
      Left            =   480
      TabIndex        =   1
      ToolTipText     =   "������J�n����ꍇ"
      Top             =   3300
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "�I��(&E)"
      Default         =   -1  'True
      Height          =   435
      Left            =   4740
      TabIndex        =   0
      ToolTipText     =   "���̍�Ƃ��I�����ă��C�����j���[�ɖ߂�ꍇ"
      Top             =   3300
      Width           =   1335
   End
   Begin ORADCLibCtl.ORADC dbcItakushaMaster 
      Height          =   315
      Left            =   2460
      Top             =   3360
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
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   "dcssvr03"
      Connect         =   "wao/wao"
      RecordSource    =   "SELECT ABITKB,ABKJNM FROM taItakushaMaster"
   End
   Begin VB.Label Label3 
      Caption         =   "�\�[�g��"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   2640
      Width           =   675
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
Attribute VB_Name = "frmKeiyakushaCheckList"
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

Private Sub cboItakusha_Click(Area As Integer)
    Select Case Area
    Case dbcAreaButton      '// 0 DB �R���{ �R���g���[����Ń{�^�����N���b�N����܂����B
    Case dbcAreaEdit        '// 1 DB �R���{ �R���g���[���̃e�L�X�g �{�b�N�X���N���b�N����܂����B
    Case dbcAreaList        '// 2 DB �R���{ �R���g���[���̃h���b�v�_�E�� ���X�g �{�b�N�X���N���b�N����܂����B
'        Debug.Print
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

Private Function pCheckDate(vDate As Variant) As Variant
    On Error GoTo pCheckDateError:
    pCheckDate = CVDate(vDate)
    Exit Function
pCheckDateError:
    Call MsgBox("�w�肳�ꂽ�Ώۓ����s���ł��B", vbCritical + vbOKOnly, mCaption)
End Function

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim StartDate As Variant
    '//Oracle �� Format �ɕϊ�����K�v������
    If "" <> Trim(txtStartDate.Text) Then
        StartDate = Format(pCheckDate(txtStartDate.Text), "YYYY/MM/DD HH:NN:SS")
        If Not IsDate(StartDate) Then
            Exit Sub
        End If
    End If
    If chkTaisho(0).Value = 0 And chkTaisho(1).Value = 0 Then
        Call MsgBox("�Ώێ҂��I������Ă��܂���B", vbCritical + vbOKOnly, mCaption)
        Exit Sub
    End If
    Dim sql As String
    sql = "SELECT * FROM ("
    sql = sql & "SELECT ABKJNM," & vbCrLf
    sql = sql & " DECODE(BAKYFG,0,NULL,1,'���','����') BAKYFGx," & vbCrLf
    '//2016/11/15 �ǉ� => ���񂹃I�[�i�[��
    sql = sql & " (select MAX(BAKJNM) from tbKeiyakushaMaster x "
    sql = sql & "   where x.BAITKB = a.BAITKB and x.BAKYCD = a.BAKYNY"
    sql = sql & " ) NAYOSENM,"
    If IsDate(StartDate) Then
        If 0 <> chkTaisho(0).Value And 0 <> chkTaisho(1).Value Then
            '//�ΏہF�V�K�o�^�� �� �ύX��
            '//�ΏہF�V�K�o�^��
            sql = sql & " CASE WHEN BAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS') THEN 1" & vbCrLf
            sql = sql & "      ELSE 0 " & vbCrLf
            sql = sql & " END as NEWCNT," & vbCrLf
            '//�ΏہF�ύX��
            sql = sql & " CASE WHEN BAADDT  < TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')       " & vbCrLf    '//�V�K�łȂ�
            sql = sql & "       AND BAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS') THEN 1" & vbCrLf    '//�C������
            sql = sql & "      ELSE 0" & vbCrLf
            sql = sql & " END as EDTCNT," & vbCrLf
        ElseIf 0 <> chkTaisho(0).Value Then
            '//�ΏہF�V�K�o�^��
            sql = sql & " CASE WHEN BAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS') THEN 1" & vbCrLf
            sql = sql & "      ELSE 0 " & vbCrLf
            sql = sql & " END as NEWCNT," & vbCrLf
            sql = sql & " 0 EDTCNT," & vbCrLf
        ElseIf 0 <> chkTaisho(1).Value Then
            '//�ΏہF�ύX��
            sql = sql & " 0 NEWCNT," & vbCrLf
            sql = sql & " CASE WHEN BAADDT  < TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')       " & vbCrLf    '//�V�K�łȂ�
            sql = sql & "       AND BAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS') THEN 1" & vbCrLf    '//�C������
            sql = sql & "      ELSE 0" & vbCrLf
            sql = sql & " END as EDTCNT," & vbCrLf
        End If
    End If
    sql = sql & " a.* " & vbCrLf
    sql = sql & " FROM tbKeiyakushaMaster a," & vbCrLf
    sql = sql & "      taItakushaMaster   b " & vbCrLf
    sql = sql & " WHERE BAITKB = ABITKB" & vbCrLf
    If -1 <> cboItakusha.BoundText Then
        sql = sql & "   AND BAITKB = " & cboItakusha.BoundText & vbCrLf
    End If
    If IsDate(StartDate) Then
        If 0 <> chkTaisho(0).Value And 0 <> chkTaisho(1).Value Then
            '//�ΏہF�V�K�o�^�� �� �ύX��
            sql = sql & "   AND(BAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
            sql = sql & "    OR BAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
            sql = sql & "   )"
        ElseIf 0 <> chkTaisho(0).Value Then
            '//�ΏہF�V�K�o�^��
            sql = sql & "   AND BAADDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
        ElseIf 0 <> chkTaisho(1).Value Then
            '//�ΏہF�ύX��
            sql = sql & "   AND BAUPDT >= TO_DATE('" & StartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
        End If
    End If
'//�o�͏���ݒ�
    Select Case cboSort.ListIndex
    Case 0      '//�_��҃J�i��
        sql = sql & " ORDER BY BAITKB,BAKNNM,BASQNO"
    Case 1      '//�X�V����
        sql = sql & " ORDER BY BAITKB,BAUPDT"
    Case Else   '//�_���
        sql = sql & " ORDER BY BAITKB,BAKYCD,BASQNO"
    End Select
    sql = sql & ")"
    
    Dim reg As New RegistryClass
    Load rptKeiyakushaCheckList
    With rptKeiyakushaCheckList
        .lblCondition.Caption = ""
        If 0 <> chkDefault.Value Then
            .lblCondition.Caption = "����F" & chkDefault.Caption
        ElseIf "" <> Trim(txtStartDate.Text) Then
            .lblCondition.Caption = "����F" & txtStartDate.Text
        End If
        .lblCondition.Caption = .lblCondition.Caption & "   �ΏێҁF"
        If 0 <> chkTaisho(0).Value And 0 <> chkTaisho(1).Value Then
            .lblCondition.Caption = .lblCondition.Caption & chkTaisho(0).Caption & "��" & chkTaisho(1).Caption
        ElseIf 0 <> chkTaisho(0).Value Then
            .lblCondition.Caption = .lblCondition.Caption & chkTaisho(0).Caption
        ElseIf 0 <> chkTaisho(1).Value Then
            .lblCondition.Caption = .lblCondition.Caption & chkTaisho(1).Caption
        End If
        .mStartDate = mStartDate
        '.mYubinCode = mYubinCode
        '.mYubinName = mYubinName
        .documentName = mCaption
'Provider=OraOLEDB.Oracle.1;Password=wao;Persist Security Info=True;User ID=wao;Data Source=dcssvr03
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    Call mForm.LockedControl(False)
    Dim sql As String, dyn As Object
    sql = "SELECT a.* FROM taSystemInformation a"
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
    
    sql = "SELECT * FROM("
    sql = sql & "SELECT '-1' ABITKB,'<< �S�Ă�Ώ� >>' ABKJNM FROM DUAL"
    sql = sql & " UNION "
    sql = sql & "SELECT ABITKB,ABKJNM FROM taItakushaMaster"
    sql = sql & ")"
    dbcItakushaMaster.RecordSource = sql
    Call dbcItakushaMaster.Refresh
    chkDefault.Value = 1
    cboSort.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFurikaeYoteiPrint = Nothing
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

