VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmHogoshaMasterRireki 
   Caption         =   "�ی�҃}�X�^���� �Ɖ�"
   ClientHeight    =   7650
   ClientLeft      =   2430
   ClientTop       =   2970
   ClientWidth     =   12750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   12750
   Begin VB.ComboBox cboFurikae 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  '�̌Œ�
      ItemData        =   "�ی�҃}�X�^�����Ɖ�.frx":0000
      Left            =   2760
      List            =   "�ی�҃}�X�^�����Ɖ�.frx":0010
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      Height          =   60
      Left            =   7140
      TabIndex        =   12
      Top             =   0
      Width           =   3975
   End
   Begin VB.Frame fraColors 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   9840
      TabIndex        =   10
      Top             =   -30
      Width           =   1215
      Begin VB.Label lblColors 
         Alignment       =   2  '��������
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "���@��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   11
         Top             =   180
         Width           =   585
      End
   End
   Begin VB.Frame fraColors 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   8520
      TabIndex        =   8
      Top             =   -30
      Width           =   1215
      Begin VB.Label lblColors 
         Alignment       =   2  '��������
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "���@��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   9
         Top             =   180
         Width           =   585
      End
   End
   Begin VB.Frame fraColors 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   7200
      TabIndex        =   6
      Top             =   -30
      Width           =   1215
      Begin VB.Label lblColors 
         Alignment       =   2  '��������
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�ʁ@��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   7
         Top             =   180
         Width           =   585
      End
   End
   Begin imText6Ctl.imText txtCAKYCD 
      Height          =   315
      Left            =   900
      TabIndex        =   0
      Top             =   120
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "�ی�҃}�X�^�����Ɖ�.frx":0034
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����Ɖ�.frx":00A2
      Key             =   "�ی�҃}�X�^�����Ɖ�.frx":00C0
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   1
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
      Format          =   "9"
      FormatMode      =   0
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   7
      LengthAsByte    =   0
      Text            =   "1234567"
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
   Begin VB.CommandButton cmdSearch 
      Caption         =   "�ΏێҌ���(&S)"
      Height          =   435
      Left            =   4980
      TabIndex        =   1
      Top             =   60
      Width           =   1300
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "�I��(&X)"
      Height          =   435
      Left            =   11100
      TabIndex        =   3
      Top             =   7020
      Width           =   1395
   End
   Begin FPSpread.vaSpread sprRireki 
      Bindings        =   "�ی�҃}�X�^�����Ɖ�.frx":0104
      Height          =   6255
      Left            =   180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   540
      Width           =   12435
      _Version        =   196608
      _ExtentX        =   21934
      _ExtentY        =   11033
      _StockProps     =   64
      ColsFrozen      =   1
      DAutoCellTypes  =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o����"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   1
      MaxRows         =   1
      OperationMode   =   1
      SpreadDesigner  =   "�ی�҃}�X�^�����Ɖ�.frx":0126
      UserResize      =   1
      VirtualMode     =   -1  'True
      VisibleCols     =   1
   End
   Begin ORADCLibCtl.ORADC dbcHogoshaMstRireki 
      Height          =   315
      Left            =   6480
      Top             =   7140
      Visible         =   0   'False
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   207
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
      DatabaseName    =   "kumon"
      Connect         =   "kumon/kumon"
      RecordSource    =   "select * from tcHogoshaMasterRireki"
   End
   Begin imDate6Ctl.imDate txtKijunBi 
      Height          =   315
      Left            =   3960
      TabIndex        =   15
      Top             =   120
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1499
      _ExtentY        =   556
      Calendar        =   "�ی�҃}�X�^�����Ɖ�.frx":03C9
      Caption         =   "�ی�҃}�X�^�����Ɖ�.frx":0549
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�ی�҃}�X�^�����Ɖ�.frx":05B7
      Keys            =   "�ی�҃}�X�^�����Ɖ�.frx":05D5
      Spin            =   "�ی�҃}�X�^�����Ɖ�.frx":0633
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
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   " "
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "2012/12/11"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   41254
      CenturyMode     =   0
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      Caption         =   "�U�֕��@"
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
      Left            =   1935
      TabIndex        =   14
      Top             =   165
      Width           =   780
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   255
      Left            =   11220
      TabIndex        =   4
      Top             =   60
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      Caption         =   "��Ű��"
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
      Left            =   180
      TabIndex        =   5
      Top             =   165
      Width           =   675
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
Attribute VB_Name = "frmHogoshaMasterRireki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mForm   As New FormClass
Private mSpread As New SpreadClass

'//2014/06/27 ���� <==> �ی�҃����e�ɔ�Ԃ̂Ńt�H�[�����e��ޔ��A�����p�ɐ錾
Private mRetForm As Form

Private Enum eFurikae
    eALL
    ePaper
    eBank
    eKaiyaku
End Enum

Private Enum eRecord
    eRireki = 0
    eMaster = 1
    eDefaultColor = 0
    eKaiyakuColor
    eRirekiColor
End Enum

Private Enum eSprCol
    eRireki = 1
    eCAHGCD = 2
    eKaiyaku = 16
    eCAKYCD = 21
End Enum

Private Sub cboFurikae_Click()
    txtKijunBi.Visible = eFurikae.ePaper = cboFurikae.ListIndex Or eFurikae.eBank = cboFurikae.ListIndex
    txtKijunBi.Value = Now()
End Sub

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Dim FieldNames As Variant
    Dim FieldIDs As Variant, IDs As Variant
    Dim ColWidths As Variant
    Dim ix As Integer
    Dim ms As New MouseClass
    
    cmdSearch.Enabled = False
    Call ms.Start
    '////////////////////////
    '//�\�����閼�O
    FieldNames = Array("R�敪", "�ی��", "�r�d�p", "�ی�Җ�", "�������`�l", "���k����", _
                       "���Z�@��", "��s��", "�x�X��", "���", "�����ԍ�", "�L��", "�ʒ��ԍ�", _
                       "�U�֊J�n", "�U�֏I��", _
                       "���", _
                       "�V�K������", _
                       "�X�V��", _
                       "�f�[�^�쐬��", "�f�[�^�X�V��", "��ŰNo" _
                )
    '////////////////////////
    '//�\�����鍀�ڂ̕ҏW
    '2012/11/15 CASQNO �� �|�P ������̂� ==> (case when length(CASQNO)=8 then casqno else null end)
    FieldIDs = Array("rKUBUN", "CAHGCD", _
            "to_char(to_date((case when length(CASQNO)=8 then casqno else null end),'yyyymmdd'),'yyyy/mm/dd')", _
                     "CAKJNM", "CAKZNM", "CASTNM", _
                     "cakkbnX", "dabknm", "dastnm", "cakzsbX", "CAKZNO", "CAYBTK", "CAYBTN", _
                     "to_char(to_date(decode(CAFKST,0,null,CAFKST),'yyyymmdd'),'yyyy/mm')", _
                     "to_char(to_date(decode(CAFKED,0,null,CAFKED),'yyyymmdd'),'yyyy/mm')", _
                     "cakyfgX", _
                     "to_char(CANWDT,'yyyy/mm/dd hh24:mi:ss')", _
                     "CAUSID", _
                     "to_char(CAADDT,'yyyy/mm/dd hh24:mi:ss')", _
                     "to_char(CAUPDT,'yyyy/mm/dd hh24:mi:ss')", "cakycd" _
                )
    ReDim ColWidths(UBound(FieldNames))
    '////////////////////////
    '//�\�������
    'defualt = 8.0
    ColWidths = Array(0, 7.6, 9.5, 14, 14, 14, 6, 12, 12, 4, 7, 3.5, 7.5, 8, 8, _
                      4, 16, 10, 16, 16, 7.6)
    sprRireki.Row = -1  '//�S�s���Ώ�
    sprRireki.MaxCols = UBound(FieldIDs) + 1
    sprRireki.ColsFrozen = 3
    For ix = LBound(FieldIDs) To UBound(FieldIDs)
        mSpread.ColWidth(ix + 1) = ColWidths(ix)
        sprRireki.Col = ix + 1      '//�w�����t�H�[�}�b�g
        Select Case FieldNames(ix)
        Case "�ی�Җ�", "���k����", "���Z�@��", "��s��", "�x�X��", "�����ԍ�", "�������`�l", "�ʒ��ԍ�", "�X�V��"
            sprRireki.TypeHAlign = TypeHAlignLeft
        Case Else
            sprRireki.TypeHAlign = TypeHAlignCenter
        End Select
    Next ix
    '////////////////////////
    '//�c�a�擾����
    IDs = Array("CAKYCD", "CAHGCD", "CASQNO", "CAKJNM", "CAKNNM", "CASTNM", _
                "cakkbn", "cabank", "casitn", "cakzsb", "CAKZNO", "CAKZNM", "CAYBTK", "CAYBTN", _
                "CAFKST", "CAFKED", _
                "cakyfg", "CANWDT", "CAUSID", "CAADDT", "CAUPDT")
    Dim sql As String
    
    On Error GoTo cmdSearch_ClickError
'    sql = "SELECT * "
'    For ix = LBound(mFieldNames) To UBound(mFieldNames)
'        sql = sql & IDs(ix) & " " & mFieldNames(ix) & ","
'    Next ix
'    sql = Left(sql, Len(sql) - 1)
'    sql = sql & " FROM tcHogoshaMasterRireki "
'    If "" <> Trim(txtCAKYCD.Text) Then
'        sql = sql & " WHERE CAKYCD = " & gdDBS.ColumnDataSet(txtCAKYCD.Text, vEnd:=True)
'    End If

    sql = "with vdBankMaster as("
    sql = sql & " select"
    sql = sql & " a.darkbn,a.dabank,a.daknnm,a.dakjnm,b.dasitn,b.daknnm dastkn,b.dakjnm dastkj,b.dasqno,b.dahtif"
    sql = sql & " from TDBANKMASTER a,TDBANKMASTER b"
    sql = sql & " Where a.dabank = b.dabank"
    sql = sql & "   and a.dasqno=':'"
    sql = sql & "   and b.dasqno='�'"  '--"�"�ȊO�͖���
    sql = sql & " order by a.dabank,b.dasitn"
    sql = sql & ")," & vbCrLf
    sql = sql & " vcHogoshaMaster as("
    sql = sql & " select a.* from tcHogoshaMaster a"
    sql = sql & " where (caitkb,cakycd,cahgcd,casqno) in("
    sql = sql & "       select caitkb,cakycd,cahgcd,max(casqno)"
    sql = sql & "       from tcHogoshaMaster "
    sql = sql & "       group by caitkb,cakycd,cahgcd"
    sql = sql & "   )"
    sql = sql & ")" & vbCrLf
    
    sql = sql & "SELECT " & vbCrLf
    For ix = LBound(FieldIDs) To UBound(FieldIDs)
        sql = sql & FieldIDs(ix) & " " & FieldNames(ix) & ","
    Next ix
    sql = Left(sql, Len(sql) - 1)
    sql = sql & " FROM(" & vbCrLf
        '///////////////////////////////
        '//�ی�҃}�X�^�[�̓��e
        '///////////////////////////////
        sql = sql & "SELECT " & vbCrLf
        For ix = LBound(IDs) To UBound(IDs)
            sql = sql & IDs(ix) & ","
        Next ix
        sql = sql & " 1 rKUBUN,SYSDATE CAMKDT," & vbCrLf
        sql = sql & " DECODE(CAKKBN,0,NULL,1,'�X�֋�','���̑�') CAKKBNx," & vbCrLf
        sql = sql & " DECODE(CAKKBN,0,DECODE(CAKZSB,1,'����',2,'����',NULL),NULL) CAKZSBx," & vbCrLf
        sql = sql & " DECODE(CAKYFG,0,NULL,1,'���','����') CAKYFGx," & vbCrLf
        sql = sql & " decode(b.DAKJNM,null,CABANK, b.DAKJNM) DABKNM," & vbCrLf
        sql = sql & " decode(b.DASTKJ,null,CASITN, b.DASTKJ) DASTNM " & vbCrLf
'//2015/02/09 �ی�҃}�X�^�̖{�̂̌����ύX����(���R�[�h�ǉ�)�ꍇ�ύX�O���o�Ȃ��̂ŕύX
       'sql = sql & " FROM vcHogoshaMaster  a," & vbCrLf
        sql = sql & " FROM tcHogoshaMaster  a," & vbCrLf
        sql = sql & "      vdBankMaster     b," & vbCrLf
        sql = sql & "      taItakushaMaster d " & vbCrLf
        sql = sql & " WHERE CABANK = b.DABANK(+)" & vbCrLf
        sql = sql & "   AND CASITN = b.DASITN(+)" & vbCrLf
        sql = sql & "   AND CAITKB = ABITKB " & vbCrLf
        If "" <> Trim(txtCAKYCD.Text) Then
'//2015/02/09 LIKE ���ɕύX
           'sql = sql & " AND CAKYCD = " & gdDBS.ColumnDataSet(txtCAKYCD.Text, vEnd:=True) & vbCrLf
            sql = sql & " AND CAKYCD LIKE " & gdDBS.ColumnDataSet("%" & txtCAKYCD.Text & "%", vEnd:=True) & vbCrLf
        End If
        Select Case cboFurikae.ListIndex
        Case eFurikae.eALL
        Case eFurikae.ePaper
            sql = sql & " and cafkst > " & Left(txtKijunBi.Number, 6) & "01" & vbCrLf
            sql = sql & " and nvl(cakyfg,'0') = '0' " & vbCrLf
        Case eFurikae.eBank
            sql = sql & " and " & Left(txtKijunBi.Number, 6) & "01" & " between cafkst and cafked " & vbCrLf
            sql = sql & " and nvl(cakyfg,'0') = '0' " & vbCrLf
        Case eFurikae.eKaiyaku
            sql = sql & " and nvl(cakyfg,'0') <> '0' " & vbCrLf
        End Select
        sql = sql & " UNION ALL " & vbCrLf
        '///////////////////////////////
        '//�ی�җ����̓��e
        '///////////////////////////////
        sql = sql & "SELECT " & vbCrLf
        For ix = LBound(IDs) To UBound(IDs)
            Select Case UCase(IDs(ix))
            Case UCase("CANWDT")
                sql = sql & " null " & IDs(ix) & ","
            Case Else
                sql = sql & IDs(ix) & ","
            End Select
        Next ix
        sql = sql & " 0 rKUBUN,CAMKDT," & vbCrLf
        sql = sql & " DECODE(CAKKBN,0,NULL,1,'�X�֋�',NULL) CAKKBNx," & vbCrLf
        sql = sql & " DECODE(CAKKBN,0,DECODE(CAKZSB,1,'����',2,'����',NULL),NULL) CAKZSBx," & vbCrLf
        sql = sql & " DECODE(CAKYFG,0,NULL,1,'���',NULL) CAKYFGx," & vbCrLf
        sql = sql & " decode(b.DAKJNM,null,CABANK, b.DAKJNM) DABKNM," & vbCrLf
        sql = sql & " decode(b.DASTKJ,null,CASITN, b.DASTKJ) DASTNM " & vbCrLf
        sql = sql & " FROM tcHogoshaMasterRireki  a," & vbCrLf
        sql = sql & "      vdBankMaster     b," & vbCrLf
        sql = sql & "      taItakushaMaster d " & vbCrLf
        sql = sql & " WHERE CABANK = b.DABANK(+)" & vbCrLf
        sql = sql & "   AND CASITN = b.DASITN(+)" & vbCrLf
        sql = sql & "   AND CAITKB = ABITKB " & vbCrLf
        If "" <> Trim(txtCAKYCD.Text) Then
'//2015/02/09 LIKE ���ɕύX
           'sql = sql & " AND CAKYCD = " & gdDBS.ColumnDataSet(txtCAKYCD.Text, vEnd:=True) & vbCrLf
            sql = sql & " AND CAKYCD LIKE " & gdDBS.ColumnDataSet("%" & txtCAKYCD.Text & "%", vEnd:=True) & vbCrLf
        End If
        If eFurikae.eALL < cboFurikae.ListIndex Then
            sql = sql & "   AND(CAKYCD,CAHGCD) in( "
            sql = sql & "   select CAKYCD,CAHGCD"
            sql = sql & "   FROM vcHogoshaMaster  a," & vbCrLf
            sql = sql & "        vdBankMaster     b," & vbCrLf
            sql = sql & "        taItakushaMaster d " & vbCrLf
            sql = sql & "   WHERE CABANK = b.DABANK(+)" & vbCrLf
            sql = sql & "     AND CASITN = b.DASITN(+)" & vbCrLf
            sql = sql & "     AND CAITKB = ABITKB " & vbCrLf
            If "" <> Trim(txtCAKYCD.Text) Then
                sql = sql & " AND CAKYCD = " & gdDBS.ColumnDataSet(txtCAKYCD.Text, vEnd:=True) & vbCrLf
            End If
            Select Case cboFurikae.ListIndex
            Case eFurikae.eALL
            Case eFurikae.ePaper
                sql = sql & " and cafkst > " & Left(txtKijunBi.Number, 6) & "01" & vbCrLf
                sql = sql & " and nvl(cakyfg,'0') = '0' " & vbCrLf
            Case eFurikae.eBank
                sql = sql & " and " & Left(txtKijunBi.Number, 6) & "01" & " between cafkst and cafked " & vbCrLf
                sql = sql & " and nvl(cakyfg,'0') = '0' " & vbCrLf
            Case eFurikae.eKaiyaku
                sql = sql & " and nvl(cakyfg,'0') <> '0' " & vbCrLf
            End Select
            sql = sql & ")" & vbCrLf
        End If
    sql = sql & ")" & vbCrLf
    'sql = sql & " ORDER BY CAKYCD,CAHGCD,CASQNO,CAMKDT DESC" & vbCrLf
    sql = sql & " ORDER BY CAKYCD,CAHGCD,CASQNO desc,rkubun desc,CAMKDT DESC" & vbCrLf
    dbcHogoshaMstRireki.RecordSource = "select * from(" & sql & ")"
    dbcHogoshaMstRireki.Refresh
    '//���z�ő�s��ݒ肵�Ȃ������Ȃ��ƃf�[�^������ɕ\������Ȃ�
    sprRireki.VirtualMaxRows = dbcHogoshaMstRireki.Recordset.RecordCount
    sprRireki.VisibleRows = sprRireki.VirtualMaxRows
    sprRireki.VirtualMode = True
    'sprRireki.OperationMode = OperationModeRow
    cmdSearch.Enabled = True
    Call sprRireki_TopLeftChange(1, 1, 1, 1)    '//�����s�̍s�J���[�ύX����������
cmdSearch_ClickError:
    cmdSearch.Enabled = True
End Sub

Private Sub Form_Activate()
    '//2014/06/27 �ی�҃}�X�^�������j��
    Unload frmHogoshaMaster
End Sub

Private Sub Form_Load()
    '//2014/06/27 �����֔�Ԃ̂Ń��j���[��ޔ�
    Set mRetForm = gdForm
    Call mForm.Init(Me, gdDBS)
    Call mSpread.Init(sprRireki)
    cboFurikae.Clear
    Call cboFurikae.AddItem("�S��", eFurikae.eALL)
    Call cboFurikae.AddItem("�U�֗p��", eFurikae.ePaper)
    Call cboFurikae.AddItem("�����U��", eFurikae.eBank)
    Call cboFurikae.AddItem("���", eFurikae.eKaiyaku)
    'cboFurikae.ItemData(eFurikae.eALL) = eFurikae.eALL
    'cboFurikae.ItemData(eFurikae.ePaper) = eFurikae.ePaper
    'cboFurikae.ItemData(eFurikae.eBank) = eFurikae.eBank
    'cboFurikae.ItemData(eFurikae.eKaiyaku) = eFurikae.eKaiyaku
    cboFurikae.ListIndex = eFurikae.eALL
    
    '//���\������ׂɃu�����N��ݒ肵�Č��������遁�O���\��
    txtCAKYCD.Text = " " '"20013"
    Call cmdSearch_Click
    txtCAKYCD.Text = ""
    sprRireki.MaxRows = 0
'    fraColors(eRecord.eDefaultColor).BackColor = RGB(255, 255, 255)
'    fraColors(eRecord.eKaiyakuColor).BackColor = RGB(255, 127, 191)
'    fraColors(eRecord.eRirekiColor).BackColor = RGB(192, 255, 239)
    Call cboFurikae_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmHogoshaMasterRireki = Nothing
    Set mForm = Nothing
    '//2014/06/27 ��������ی�҃����e�ɔ�Ԃ̂Ń��j���[�ɕ���
    Set gdForm = mRetForm
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

Private Sub sprRireki_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Col < 1 And Row < 1 Then
        Exit Sub
    End If
    Dim frm As Form
    Set frm = frmHogoshaMaster
    Call frm.Show
    frm.txtCAKYCD.Text = mSpread.Text(eSprCol.eCAKYCD, Row)
    frm.txtCAHGCD.Text = mSpread.Text(eSprCol.eCAHGCD, Row)
    Call frm.txtCAHGCD_KeyDown(vbKeyReturn, 0)
    Set gdForm = Me
End Sub

Private Sub sprRireki_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    Dim Row As Long, data As Variant
    'sprRireki.BlockMode = True
    For Row = NewTop To NewTop + 24
        If Row <= mSpread.MaxRows Then
            mSpread.BackColor(-1, Row) = fraColors(eRecord.eDefaultColor).BackColor
            '//�������H
            If eRecord.eMaster <> mSpread.Text(eSprCol.eRireki, Row) Then
                mSpread.BackColor(-1, Row) = fraColors(eRecord.eRirekiColor).BackColor
            Else
                '//����ԁH
                If "" <> mSpread.Text(eSprCol.eKaiyaku, Row) Then
                    mSpread.BackColor(-1, Row) = fraColors(eRecord.eKaiyakuColor).BackColor
                End If
            End If
        End If
    Next Row
    'sprRireki.BlockMode = False
End Sub

Private Sub txtCAKYCD_KeyDown(KeyCode As Integer, Shift As Integer)
    '// Return �܂��� Shift�{TAB �̂Ƃ��̂ݏ�������
    If Not (KeyCode = vbKeyReturn) Then
        Exit Sub
    ElseIf 0 = Len(Trim(txtCAKYCD.Text)) Then
        Exit Sub
    End If
'//2013/06/18 �O�[�����ߍ���
    txtCAKYCD.Text = Format(Val(txtCAKYCD.Text), String(7, "0"))
End Sub
