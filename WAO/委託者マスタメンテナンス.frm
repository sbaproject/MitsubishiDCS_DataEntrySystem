VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Begin VB.Form frmItakushaMaster 
   Caption         =   "�ϑ��҃}�X�^�����e�i���X"
   ClientHeight    =   4845
   ClientLeft      =   2730
   ClientTop       =   2235
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6990
   Begin VB.ComboBox cboABDEFF 
      Height          =   300
      ItemData        =   "�ϑ��҃}�X�^�����e�i���X.frx":0000
      Left            =   2340
      List            =   "�ϑ��҃}�X�^�����e�i���X.frx":000A
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   15
      Top             =   1980
      Width           =   975
   End
   Begin VB.Frame fraShoriKubun 
      Caption         =   "�����敪"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Tag             =   "InputKey"
      Top             =   120
      Width           =   3915
      Begin VB.OptionButton optShoriKubun 
         Caption         =   "�Q��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2940
         TabIndex        =   20
         Tag             =   "InputKey"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optShoriKubun 
         Caption         =   "�C��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   7
         Tag             =   "InputKey"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optShoriKubun 
         Caption         =   "�폜"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   6
         Tag             =   "InputKey"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optShoriKubun 
         Caption         =   "�V�K"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Tag             =   "InputKey"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblShoriKubun 
         BackColor       =   &H000000FF&
         Caption         =   "�����敪"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "���~(&C)"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2700
      TabIndex        =   1
      Top             =   4080
      Width           =   1395
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�X�V(&U)"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   900
      TabIndex        =   0
      Top             =   4080
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "�I��(&X)"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4920
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin ORADCLibCtl.ORADC dbcItakushaMaster 
      Height          =   315
      Left            =   2940
      Top             =   3600
      Visible         =   0   'False
      Width           =   1875
      _Version        =   65536
      _ExtentX        =   3307
      _ExtentY        =   556
      _StockProps     =   207
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
      RecordSource    =   "SELECT * FROM taItakushaMaster"
   End
   Begin imText6Ctl.imText txtABITCD 
      Height          =   285
      Left            =   2340
      TabIndex        =   9
      Tag             =   "InputKey"
      Top             =   1140
      Width           =   615
      _Version        =   65537
      _ExtentX        =   1085
      _ExtentY        =   503
      Caption         =   "�ϑ��҃}�X�^�����e�i���X.frx":001C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ϑ��҃}�X�^�����e�i���X.frx":0088
      Key             =   "�ϑ��҃}�X�^�����e�i���X.frx":00A6
      MouseIcon       =   "�ϑ��҃}�X�^�����e�i���X.frx":00EA
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
      MaxLength       =   5
      LengthAsByte    =   0
      Text            =   "12345"
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
   Begin imText6Ctl.imText txtABKJNM 
      DataField       =   "ABKJNM"
      DataSource      =   "dbcItakushaMaster"
      Height          =   285
      Left            =   2340
      TabIndex        =   12
      Top             =   1560
      Width           =   1575
      _Version        =   65537
      _ExtentX        =   2778
      _ExtentY        =   503
      Caption         =   "�ϑ��҃}�X�^�����e�i���X.frx":0106
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ϑ��҃}�X�^�����e�i���X.frx":0172
      Key             =   "�ϑ��҃}�X�^�����e�i���X.frx":0190
      MouseIcon       =   "�ϑ��҃}�X�^�����e�i���X.frx":01D4
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
      MaxLength       =   16
      LengthAsByte    =   0
      Text            =   "���������D�D�D��"
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   4
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.Label lblABITKB 
      BackColor       =   &H000000FF&
      Caption         =   "�ϑ��ҋ敪"
      DataField       =   "ABITKB"
      DataSource      =   "dbcItakushaMaster"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3540
      TabIndex        =   19
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblABDEFF 
      BackColor       =   &H000000FF&
      Caption         =   "�f�t�H���g"
      DataField       =   "ABDEFF"
      DataSource      =   "dbcItakushaMaster"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3540
      TabIndex        =   18
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblABUSID 
      BackColor       =   &H000000FF&
      Caption         =   "�X�V��"
      DataField       =   "ABUSID"
      DataSource      =   "dbcItakushaMaster"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5580
      TabIndex        =   17
      Top             =   2940
      Width           =   975
   End
   Begin VB.Label lblABUPDT 
      BackColor       =   &H000000FF&
      Caption         =   "�X�V��"
      DataField       =   "ABUPDT"
      DataSource      =   "dbcItakushaMaster"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5580
      TabIndex        =   16
      Top             =   3300
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  '�E����
      Caption         =   "�f�t�H���g�\��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   900
      TabIndex        =   14
      Top             =   1980
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      Caption         =   "�ϑ��Җ�(����)"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   900
      TabIndex        =   13
      Top             =   1605
      Width           =   1275
   End
   Begin VB.Label lblABITCD 
      BackColor       =   &H000000FF&
      Caption         =   "�ϑ��Ҕԍ�"
      DataField       =   "ABITCD"
      DataSource      =   "dbcItakushaMaster"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3540
      TabIndex        =   11
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
      Caption         =   "�ϑ��Ҕԍ�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   900
      TabIndex        =   10
      Tag             =   "InputKey"
      Top             =   1140
      Width           =   1275
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4980
      TabIndex        =   3
      Top             =   60
      Width           =   1395
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
Attribute VB_Name = "frmItakushaMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm As New FormClass
Private mCaption As String

Private Sub pLockedControl(blMode As Boolean)
    Call mForm.LockedControl(blMode)
    cmdEnd.Enabled = blMode
End Sub

Private Sub cboABDEFF_Click()
    lblABDEFF.Caption = Abs(cboABDEFF.ListIndex > 0)
End Sub

Private Sub cmdUpdate_Click()
'    Dim sql As String, dyn As OraDynaset
    Dim sql As String, dyn As Object
    If lblShoriKubun.Caption = eShoriKubun.Delete Then
        sql = "SELECT COUNT(*) AS CNT FROM tbkeiyakushaMaster"
        sql = sql & " WHERE BAITKB = '" & lblABITKB.Caption & "'"
#If ORA_DEBUG = 1 Then
        Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
        Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
        If Val(gdDBS.Nz(dyn.Fields("CNT"))) Then
            Call MsgBox("�_��҃}�X�^�Ŏg�p����Ă��邽��" & vbCrLf & vbCrLf & "�폜���鎖�͏o���܂���.", vbCritical, mCaption)
            Exit Sub
        End If
        If vbOK <> MsgBox("�폜���܂����H" & vbCrLf & vbCrLf & "���ɖ߂����Ƃ͏o���܂���.", vbInformation + vbOKCancel + vbDefaultButton2, mCaption) Then
            Exit Sub
        Else
'//2002/11/26 OIP-00000 ORA-04108 �ŃG���[�ɂȂ�̂� Execute() �Ŏ��s����悤�ɕύX.
'// Oracle Data Control 8i(3.6) 9i(4.2) �̈Ⴂ���ȁH
'//            Call dbcItakushaMaster.Recordset.Delete
            Call dbcItakushaMaster.UpdateControls
            sql = "DELETE taItakushaMaster"
            sql = sql & " WHERE ABITCD = '" & lblABITCD.Caption & "'"
            Call gdDBS.Database.ExecuteSQL(sql)
        End If
    Else
        If Not IsNumeric(lblABITKB.Caption) Then
            sql = "SELECT MAX(ABITKB) AS MaxCode FROM taItakushaMaster"
#If ORA_DEBUG = 1 Then
            Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
            Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
            If IsNull(dyn.Fields("MaxCode")) Then
                lblABITKB.Caption = 0
            Else
                lblABITKB.Caption = Val(gdDBS.Nz(dyn.Fields("MaxCode"))) + 1
            End If
            Call dyn.Close
        End If
        lblABUSID.Caption = gdDBS.LoginUserName
        lblABUPDT.Caption = gdDBS.sysDate
        Call dbcItakushaMaster.UpdateRecord
    End If
    Call pLockedControl(True)
    Call txtABITCD.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Call dbcItakushaMaster.UpdateControls
    Call pLockedControl(True)
    Call txtABITCD.SetFocus
End Sub

Private Sub cmdEnd_Click()
    Call dbcItakushaMaster.UpdateControls
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    dbcItakushaMaster.RecordSource = ""
    Call pLockedControl(True)
    Call mForm.pInitControl
    '//�����l���Z�b�g�F�Q�ƃ��[�h
    optShoriKubun(eShoriKubun.Refer).Value = True
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmItakushaMaster = Nothing
    Set mForm = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub lblABDEFF_Change()
    If IsNumeric(lblABDEFF.Caption) Then
        cboABDEFF.ListIndex = Val(lblABDEFF.Caption)
    End If
End Sub

Private Sub optShoriKubun_Click(Index As Integer)
    On Error Resume Next    'Form_Load()���Ƀt�H�[�J�X�𓖂Ă��Ȃ����G���[�ƂȂ�̂ŉ���̃G���[����
    lblShoriKubun.Caption = Index
    Call txtABITCD.SetFocus
End Sub

Private Sub txtABITCD_KeyDown(KeyCode As Integer, Shift As Integer)
    '// Return �̂Ƃ��̂ݏ�������
    If Not (KeyCode = vbKeyReturn) Then
        Exit Sub
    End If
'    Dim sql As String, dyn As OraDynaset
    Dim sql As String, dyn As Object
    Dim msg As String
    sql = "SELECT * FROM taItakushaMaster"
    sql = sql & " WHERE ABITCD = '" & txtABITCD.Text & "'"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If 0 = dyn.RecordCount Then
        If eShoriKubun.Add <> lblShoriKubun.Caption Then     '���R�[�h�����ŐV�K�ȊO�̎�
            msg = "�Y���f�[�^�͑��݂��܂���."
        End If
    ElseIf eShoriKubun.Add = lblShoriKubun.Caption Then      '���R�[�h�L��ŐV�K�̎�
        msg = "���Ƀf�[�^�����݂��܂�."
    End If
    Set dyn = Nothing
    If msg <> "" Then
        Call MsgBox(msg, vbInformation, mCaption)
        Call txtABITCD.SetFocus
        Exit Sub
    End If
    dbcItakushaMaster.RecordSource = sql
    Call dbcItakushaMaster.Refresh
    If dbcItakushaMaster.Recordset.RecordCount = 0 Then
        Call dbcItakushaMaster.Recordset.AddNew
        lblABITCD.Caption = txtABITCD.Text
        lblABDEFF.Caption = 0
    Else
        Call dbcItakushaMaster.Recordset.MoveFirst
        Call dbcItakushaMaster.Recordset.Edit
    End If
    '//�Q�ƂŖ�����΃{�^���̐���J�n
    If False = optShoriKubun(eShoriKubun.Refer).Value Then
        Call pLockedControl(False)
    End If
    '//�R���g���[���������ԍ��ɂ����������߂ɂ��܂��Ȃ��F���ɕ��@��������Ȃ��H
    Call SendKeys("+{TAB}+{TAB}")
End Sub

Private Sub mnuEnd_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

