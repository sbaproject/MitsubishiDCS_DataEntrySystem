VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmFurikaeIraishoPrint 
   Caption         =   "�����U�ֈ˗���(���)"
   ClientHeight    =   3975
   ClientLeft      =   3975
   ClientTop       =   3135
   ClientWidth     =   6330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6330
   Begin VB.ComboBox cboPrintUser 
      Height          =   300
      Left            =   1920
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   5
      Top             =   1680
      Width           =   2115
   End
   Begin VB.ComboBox cboPrintDate 
      Height          =   300
      Left            =   1920
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   4
      Top             =   1260
      Width           =   2115
   End
   Begin ORADCLibCtl.ORADC dbcItakushaMaster 
      Height          =   315
      Left            =   2100
      Top             =   3360
      Visible         =   0   'False
      Width           =   1995
      _Version        =   65536
      _ExtentX        =   3519
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
      Connect         =   "kumon/kumon"
      RecordSource    =   "SELECT ABITKB,ABKJNM FROM taItakushaMaster"
      ReadOnly        =   -1  'True
   End
   Begin VB.Frame fraSort 
      Caption         =   "�o�͏���"
      Height          =   915
      Left            =   1860
      TabIndex        =   6
      Tag             =   "0"
      Top             =   2160
      Width           =   1875
      Begin VB.OptionButton optSort 
         Caption         =   "�f�[�^���� ��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   540
         Width           =   1575
      End
      Begin VB.OptionButton optSort 
         Caption         =   "�_��Ҕԍ� ��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraImportX 
      BackColor       =   &H000000FF&
      Caption         =   "�Ώێ�(�捞��)"
      Height          =   1035
      Left            =   5280
      TabIndex        =   16
      Top             =   2040
      Width           =   1695
      Begin VB.CheckBox chkTaishoX 
         Caption         =   "�V�K�o�^��"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkTaishoX 
         Caption         =   "�C����"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame fraInputX 
      BackColor       =   &H000000FF&
      Caption         =   "�Ώێ�(����͕�)"
      Height          =   1035
      Left            =   5280
      TabIndex        =   13
      Top             =   900
      Width           =   1695
      Begin VB.CheckBox chkTaishoX 
         Caption         =   "�C����"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Top             =   600
         Value           =   1  '����
         Width           =   1335
      End
      Begin VB.CheckBox chkTaishoX 
         Caption         =   "�V�K�o�^��"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Value           =   1  '����
         Width           =   1455
      End
   End
   Begin MSDBCtls.DBCombo cboItakusha 
      Bindings        =   "�����U�ֈ˗���(���).frx":0000
      Height          =   300
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   529
      _Version        =   393216
      Style           =   2
      ListField       =   "ABKJNM"
      BoundColumn     =   "ABITKB"
      Text            =   "�ϑ��҈ꗗ"
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   1875
   End
   Begin imText6Ctl.imText txtStartDate 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
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
      Left            =   360
      TabIndex        =   10
      ToolTipText     =   "������J�n����ꍇ"
      Top             =   3300
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "�I��(&E)"
      Height          =   435
      Left            =   4620
      TabIndex        =   0
      ToolTipText     =   "���̍�Ƃ��I�����ă��C�����j���[�ɖ߂�ꍇ"
      Top             =   3300
      Width           =   1335
   End
   Begin VB.Label lblPrintUser 
      Alignment       =   1  '�E����
      Caption         =   "�S����"
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   1740
      Width           =   615
   End
   Begin VB.Label lblPrintDate 
      Alignment       =   1  '�E����
      Caption         =   "�o�͓�"
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
      Caption         =   "�ϑ���"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   900
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label1"
      Height          =   195
      Left            =   4860
      TabIndex        =   9
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
Attribute VB_Name = "frmFurikaeIraishoPrint"
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
Private Const pcALL_USER As String = "<< �S�Ă�Ώ� >>"

Private Enum eSort
    eKeiyakusha = 0
    eInput
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

Private Sub cboPrintDate_Click()
    Me.Refresh
End Sub

Private Sub chkDefault_Click()
    If 0 = chkDefault.Value Then
        txtStartDate.Enabled = True
    Else
        txtStartDate.Text = mStartDate
        txtStartDate.Enabled = False
        txtStartDate_LostFocus      '//�o�͓��E�S���҂����t���b�V��
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

Private Function pPrintDateUpdate(ByVal vUpdDate As Variant, ByVal vPrnDate As Variant) As Variant
    Dim prnDate As Date
    
    If True = IsDate(vPrnDate) Then
        '//�Ώۂ̏o�͓����I������Ă���ꍇ�͂��̒l���̗p���ďI��
        pPrintDateUpdate = vPrnDate
        Exit Function
    End If
    prnDate = gdDBS.sysDate()
    Dim sql As String, where As String, updCnt As Long
    
    updCnt = 0
    On Error GoTo pPrintDateUpdateError
    '//�g�����U�N�V��������������
    Call gdDBS.Database.BeginTrans
    Call gdDBS.TriggerControl("tcHogoshaMaster", False)
    
    '//�ی�҃}�X�^�̃��C�� WHERE ��𐶐�
    where = " WHERE CACHEK IS NULL "
    where = where & "   AND CAUPDT >= TO_DATE('" & vUpdDate & "','YYYY/MM/DD HH24:MI:SS')"
    If cboPrintUser.Text <> pcALL_USER Then
        where = where & "   AND UPPER(CAUSID) = '" & UCase(cboPrintUser.Text) & "'"
    End If
    '////////////////////////
    '//��ɗ������X�V
    sql = "UPDATE tcHogoshaMasterRireki     SET "
    sql = sql & " CACHEK = TO_DATE('" & prnDate & "','YYYY/MM/DD HH24:MI:SS')"
    sql = sql & " WHERE CACHEK IS NULL "
    '//�����R�[�h(CAKSCD)�̓L�[���ڂł͂Ȃ��FPrimaryKey �ݒ肳��Ă͂��邪�I�I�I
    sql = sql & "   AND (CAITKB,CAKYCD,CAHGCD,CASQNO) IN ("
    sql = sql & " SELECT CAITKB,CAKYCD,CAHGCD,CASQNO"
    sql = sql & " FROM tcHogoshaMaster "
    sql = sql & where
    sql = sql & " )"
    updCnt = updCnt + gdDBS.Database.ExecuteSQL(sql)
    '////////////////////////
    '//�}�X�^���X�V
    sql = "UPDATE tcHogoshaMaster           SET "
    sql = sql & " CACHEK = TO_DATE('" & prnDate & "','YYYY/MM/DD HH24:MI:SS')"
    sql = sql & where
    updCnt = updCnt + gdDBS.Database.ExecuteSQL(sql)
    Call gdDBS.Database.CommitTrans
    Call gdDBS.TriggerControl("tcHogoshaMaster")
    If 0 = updCnt Then
        '//�V�K�������������ꍇ�o�͂��Ȃ��悤�ɂ���
        Call MsgBox("�����s���̃f�[�^�͑��݂��܂���ł����B", vbCritical + vbOKOnly, mCaption)
    Else
        '//����ɏ������ꂽ�̂ōX�V����ԋp
        pPrintDateUpdate = prnDate
        '//�o�͓��̃��t���b�V��
        Call pPrintDateRefresh
        cboPrintDate.ListIndex = 1  '//�ŐV�� Index = 1 �̂͂�
    End If
    Exit Function
pPrintDateUpdateError:
    Call gdDBS.Database.Rollback
    Call gdDBS.TriggerControl("tcHogoshaMaster")
    Call gdDBS.ErrorCheck(gdDBS.Database)
End Function

Private Sub cmdPrint_Click()
    Dim StartDate As Variant
    '//Oracle �� Format �ɕϊ�����K�v������
    If "" = Trim(txtStartDate.Text) Then
        Call MsgBox("����͕K�{���͍��ڂł��B", vbCritical + vbOKOnly, mCaption)
        Exit Sub
    Else
        StartDate = Format(pCheckDate(txtStartDate.Text), "YYYY/MM/DD HH:NN:SS")
        If Not IsDate(StartDate) Then
            Call MsgBox("�w�肳�ꂽ������s���ł��B", vbCritical + vbOKOnly, mCaption)
            Exit Sub
        End If
    End If
    
    Dim ms As New MouseClass
    Call ms.Start

    Dim sql As String
    Dim prnDate As Variant
    '//�����s���Ƃ����ꍇ�A���݂̖����s����S����ɍX�V����B
    prnDate = pPrintDateUpdate(txtStartDate.Text, cboPrintDate.Text)
    If True = IsEmpty(prnDate) Then
        Exit Sub
    End If
    
    Dim Field As Variant, ix As Integer
    Field = Array("CAITKB", "CAKYCD", "CAKSCD", "CAHGCD", "CASQNO", "CAKJNM", "CAKNNM", "CASTNM", "CAKKBN", _
                  "CABANK", "CASITN", "CAKZSB", "CAKZNO", "CAYBTK", "CAYBTN", "CAKZNM", "CAKYST", "CAKYED", _
                  "CAFKST", "CAFKED", "CASKGK", "CAHKGK", "CAKYDT", "CAKYFG", "CATRFG", "CAUSID", "CAADDT", _
                  "CAUPDT", "CACHEK")
    sql = "SELECT * FROM(" & vbCrLf
        
        '///////////////////////////////
        '//�ی�҃}�X�^�[�̓��e
        '///////////////////////////////
        sql = sql & "SELECT 1 rKUBUN,SYSDATE CAMKDT," & vbCrLf
        For ix = LBound(Field) To UBound(Field)
            sql = sql & Field(ix) & ","
        Next ix
        sql = sql & " DECODE(CAKKBN,0,NULL,1,'�X','��') CAKKBNx," & vbCrLf
        sql = sql & " DECODE(CAKKBN,0,DECODE(CAKZSB,1,'��',2,'��','��'),NULL) CAKZSBx," & vbCrLf
        sql = sql & " DECODE(CAKYFG,0,NULL,1,'���','����') CAKYFGx," & vbCrLf
        sql = sql & " b.DAKJNM BankName," & vbCrLf
        sql = sql & " c.DAKJNM ShitenName," & vbCrLf
        sql = sql & " d.ABKJNM, " & vbCrLf
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
        sql = sql & " AND CACHEK = TO_DATE('" & prnDate & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf
        '//���[�U�[�w��̏ꍇ
        If 0 < cboPrintUser.ListIndex Then
            sql = sql & " AND UPPER(CAUSID) = '" & UCase(cboPrintUser.Text) & "'" & vbCrLf
        End If
        sql = sql & " UNION ALL " & vbCrLf
        '///////////////////////////////
        '//�ی�җ����̓��e
        '///////////////////////////////
        sql = sql & "SELECT 0 rKUBUN,CAMKDT," & vbCrLf
        For ix = LBound(Field) To UBound(Field)
            sql = sql & Field(ix) & ","
        Next ix
        sql = sql & " DECODE(CAKKBN,0,NULL,1,'�X',NULL) CAKKBNx," & vbCrLf
        sql = sql & " DECODE(CAKKBN,0,DECODE(CAKZSB,1,'��',2,'��',NULL),NULL) CAKZSBx," & vbCrLf
        sql = sql & " DECODE(CAKYFG,0,NULL,1,'���',NULL) CAKYFGx," & vbCrLf
        sql = sql & " b.DAKJNM BankName," & vbCrLf
        sql = sql & " c.DAKJNM ShitenName," & vbCrLf
        sql = sql & " d.ABKJNM," & vbCrLf
        sql = sql & " (" & vbCrLf
        sql = sql & "   SELECT "
'//2008/07/17 ���ʃZ�b�g���Q���Ԃ�̂ŃG���[�ɂȂ菈���ł��Ȃ��Ȃ����̂� MAX() �ɕύX   / 2012/09/20 �����𕜊������̂� MAX() ����
        sql = sql & "   x.CAUPDT " & vbCrLf     '//���͓�-Sort �p
        sql = sql & "   FROM tcHogoshaMaster  x " & vbCrLf
        sql = sql & "   WHERE x.CAITKB = a.CAITKB " & vbCrLf
        sql = sql & "     AND x.CAKYCD = a.CAKYCD " & vbCrLf
        sql = sql & "     AND x.CAKSCD = a.CAKSCD " & vbCrLf   '//�����R�[�h�͏��O / 2012/09/20 ����
        sql = sql & "     AND x.CAHGCD = a.CAHGCD " & vbCrLf
        sql = sql & "     AND x.CASQNO = a.CASQNO " & vbCrLf
        sql = sql & " ) INPDATE," & vbCrLf
        sql = sql & " (" & vbCrLf
        sql = sql & "   SELECT "
'//2008/07/17 ���ʃZ�b�g���Q���Ԃ�̂ŃG���[�ɂȂ菈���ł��Ȃ��Ȃ����̂� MAX() �ɕύX   / 2012/09/20 �����𕜊������̂� MAX() ����
        sql = sql & "   x.CAUSID " & vbCrLf     '//���̓��[�U�[-Sort �p
        sql = sql & "   FROM tcHogoshaMaster  x " & vbCrLf
        sql = sql & "   WHERE x.CAITKB = a.CAITKB " & vbCrLf
        sql = sql & "     AND x.CAKYCD = a.CAKYCD " & vbCrLf
        sql = sql & "     AND x.CAKSCD = a.CAKSCD " & vbCrLf   '//�����R�[�h�͏��O / 2012/09/20 ����
        sql = sql & "     AND x.CAHGCD = a.CAHGCD " & vbCrLf
        sql = sql & "     AND x.CASQNO = a.CASQNO " & vbCrLf
        sql = sql & " )  INPUSER " & vbCrLf
        sql = sql & " FROM tcHogoshaMasterrireki  a," & vbCrLf
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
        sql = sql & " AND CACHEK = TO_DATE('" & prnDate & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf
        '//���[�U�[�w��̏ꍇ
        If 0 < cboPrintUser.ListIndex Then
            sql = sql & " AND (CAITKB,CAKYCD,CAHGCD,CASQNO) IN (" & vbCrLf
            sql = sql & "   SELECT CAITKB,CAKYCD,CAHGCD,CASQNO" & vbCrLf
            sql = sql & "   FROM tcHogoshaMaster " & vbCrLf
            sql = sql & "   WHERE 1 = 1" & vbCrLf
            If -1 <> cboItakusha.BoundText Then
                sql = sql & "   AND CAITKB = " & cboItakusha.BoundText & vbCrLf
            End If
            sql = sql & "     AND CACHEK = TO_DATE('" & prnDate & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf
            sql = sql & "     AND UPPER(CAUSID) = '" & UCase(cboPrintUser.Text) & "'" & vbCrLf
            sql = sql & "   )" & vbCrLf
        End If
    sql = sql & ")" & vbCrLf
#If 0 Then
'ORDER BY --INPUSER,
'    CAITKB,CAKYCD,CAHGCD,cakscd desc,CASQNO desc,CAMKDT DESC
    Select Case Val(fraSort.Tag)
    Case eSort.eKeiyakusha
        sql = sql & " ORDER BY " & vbCrLf
    Case eSort.eInput
        sql = sql & " ORDER BY INPUSER,INPDATE desc," & vbCrLf
    End Select
    sql = sql & "CAITKB,CAKYCD,CAHGCD,cakscd desc,CASQNO desc,CAMKDT DESC" & vbCrLf
#Else
'//2012/10/16 �^�p�ケ�̎��Ŗ����Ƃ����Ȃ��F�������j���
    Select Case Val(fraSort.Tag)
    Case eSort.eKeiyakusha
        sql = sql & " ORDER BY INPUSER,CAITKB,CAKYCD,CAHGCD,CASQNO,CAMKDT DESC" & vbCrLf
    Case eSort.eInput
        sql = sql & " ORDER BY INPUSER,INPDATE,CAITKB,CAKYCD,CAHGCD,CASQNO,CAMKDT DESC" & vbCrLf
    End Select
#End If
    Dim reg As New RegistryClass
    Load rptKouzaFurikaeIraisho
    With rptKouzaFurikaeIraisho
        If 0 <> chkDefault.Value Then
            .lblCondition.Caption = "����F" & chkDefault.Caption '//�u�O��ݐϓ��v���\�������
        ElseIf "" <> Trim(txtStartDate.Text) Then
            .lblCondition.Caption = "����F[" & txtStartDate.Text & "]"
        End If
        .lblCondition.Caption = .lblCondition.Caption & "  �o�͓��F[" & cboPrintDate.Text & "]  �S���ҁF" & cboPrintUser.Text
        .lblCondition.Caption = .lblCondition.Caption & "  �o�͏��ԁF" & optSort(Val(fraSort.Tag)).Caption
        .mStartDate = txtStartDate.Text
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
    Set ms = Nothing
End Sub

Private Sub Form_Activate()
    If "" = Trim(cboItakusha.BoundText) Then
        cboItakusha.BoundText = "-1"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mForm.KeyDown(KeyCode, Shift)
End Sub

Private Sub pPrintDateRefresh()
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    '////////////////////////////////////
    '//�������Ώۓ��t
    sql = "SELECT DISTINCT "
    sql = sql & " 0 AS KEY,"
    sql = sql & " 0 AS SEQ,"
    sql = sql & " '<< �����s�� >>' cachek "
    sql = sql & " FROM DUAL "
    sql = sql & " UNION "
    sql = sql & " SELECT "
    sql = sql & " 1 AS KEY,"
    sql = sql & " ROWNUM SEQ,"
    sql = sql & " cachek "
    sql = sql & " FROM("
    sql = sql & "   SELECT DISTINCT "
    sql = sql & "   TO_CHAR(CACHEK,'yyyy/mm/dd hh24:mi:ss') cachek "
    sql = sql & "   FROM tcHogoshaMaster "
    sql = sql & "   WHERE CACHEK >= TO_DATE('" & txtStartDate.Text & "','yyyy/mm/dd hh24:mi:ss')"
    sql = sql & " )"
    sql = sql & " ORDER BY KEY,CACHEK DESC"
    Set dyn = gdDBS.OpenRecordset(sql)
    Dim idx As Integer, txt As String
    idx = cboPrintDate.ListIndex
    txt = cboPrintDate.Text
    cboPrintDate.Clear
    Do Until dyn.EOF
        Call cboPrintDate.AddItem(dyn.Fields("cachek").Value)
        If txt = cboPrintDate.List(cboPrintDate.NewIndex) Then
            idx = cboPrintDate.NewIndex
        End If
        Call dyn.MoveNext
    Loop
    dyn.Close
    If idx = -1 Then
        idx = 0
    End If
    cboPrintDate.ListIndex = idx
End Sub

Private Sub pPrintUserRefresh(Optional vUserName As String = "")
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    '////////////////////////////////////
    '//�������Ώۃ��[�U�[
    sql = "SELECT DISTINCT "
    sql = sql & " 0 AS KEY,"
    sql = sql & " 0 AS SEQ,"
    sql = sql & " '" & pcALL_USER & "' causid "
    sql = sql & " FROM DUAL "
    sql = sql & " UNION "
    sql = sql & " SELECT "
    sql = sql & " 1 AS KEY,"
    sql = sql & " ROWNUM AS SEQ,"
    sql = sql & " causid "
    sql = sql & " FROM("
    sql = sql & "   SELECT "
    sql = sql & "   '" & gdDBS.LoginUserName & "' causid "
    sql = sql & "   FROM DUAL "
    sql = sql & "   UNION "
    sql = sql & "   SELECT DISTINCT "
    'sql = sql & "   UPPER(causid) causid "
    sql = sql & "   causid "
    sql = sql & "   FROM tcHogoshaMaster "
    sql = sql & "   WHERE cachek >= TO_DATE('" & txtStartDate.Text & "','yyyy/mm/dd hh24:mi:ss')"
    sql = sql & " )"
    sql = sql & " ORDER BY KEY,causid"
    Set dyn = gdDBS.OpenRecordset(sql)
    Dim idx As Integer, txt As String
    idx = cboPrintUser.ListIndex
    If vUserName = "" Then
        txt = cboPrintUser.Text
    Else
        txt = vUserName     '//�����\�����[�U�[�����O�C�����[�U�[�ɂ���
    End If
    cboPrintUser.Clear
    Do Until dyn.EOF
        Call cboPrintUser.AddItem(dyn.Fields("causid").Value)
        If txt = cboPrintUser.List(cboPrintUser.NewIndex) Then
            idx = cboPrintUser.NewIndex
        End If
        Call dyn.MoveNext
    Loop
    dyn.Close
    If idx = -1 Then
        idx = 0
    End If
    cboPrintUser.ListIndex = idx
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    Call mForm.LockedControl(False)
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
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
    
    Call pPrintDateRefresh
    Call pPrintUserRefresh(gdDBS.LoginUserName())
    
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
    Set frmFurikaeIraishoPrint = Nothing
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

Private Sub optSort_Click(Index As Integer)
    fraSort.Tag = Index
End Sub

Private Sub txtStartDate_LostFocus()
    Call pPrintDateRefresh
    Call pPrintUserRefresh
End Sub
