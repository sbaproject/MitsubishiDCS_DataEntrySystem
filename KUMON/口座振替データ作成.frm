VERSION 5.00
Object = "{73B346C1-F158-11D1-AF40-006097476B29}#1.0#0"; "Date60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmYoteiDataExport 
   Caption         =   "�����U�փf�[�^�쐬"
   ClientHeight    =   4470
   ClientLeft      =   2295
   ClientTop       =   2235
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6765
   Begin MSComctlLib.ProgressBar pgbRecord 
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   3300
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdMakeText 
      Caption         =   "�e�L�X�g�쐬(&T)"
      Height          =   435
      Left            =   1620
      TabIndex        =   3
      Top             =   3840
      Width           =   1395
   End
   Begin VB.ComboBox cboFurikaeBi 
      Height          =   300
      Left            =   3060
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   310
      Width           =   1275
   End
   Begin VB.CheckBox chkJisseki 
      BackColor       =   &H000000FF&
      Caption         =   "1 = �m��"
      Height          =   315
      Left            =   180
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Value           =   2  '����
      Width           =   975
   End
   Begin VB.CommandButton cmdOutMsg 
      Caption         =   "�쐬����(&L)"
      Height          =   435
      Left            =   3240
      TabIndex        =   4
      Top             =   3840
      Width           =   1395
   End
   Begin VB.CommandButton cmdMakeDB 
      Caption         =   "�c�a�쐬(&D)"
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "�I��(&X)"
      Height          =   435
      Left            =   5280
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin imDate6Ctl.imDate txtFurikaeBi 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1035
      _Version        =   65537
      _ExtentX        =   1826
      _ExtentY        =   503
      Calendar        =   "�����U�փf�[�^�쐬.frx":0000
      Caption         =   "�����U�փf�[�^�쐬.frx":0186
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "�����U�փf�[�^�쐬.frx":01F4
      Keys            =   "�����U�փf�[�^�쐬.frx":0212
      MouseIcon       =   "�����U�փf�[�^�쐬.frx":0270
      Spin            =   "�����U�փf�[�^�쐬.frx":028C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   255
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
      MaxDate         =   73050
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
      Text            =   "    /  /  "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   -2
      CenturyMode     =   0
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   4740
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label1"
      Height          =   315
      Left            =   5220
      TabIndex        =   9
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label8 
      Caption         =   "�����U�֓�"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   360
      Width           =   915
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "�l�r ����"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   5895
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
Attribute VB_Name = "frmYoteiDataExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCaption As String
Private Const mExeMsg As String = "��Ǝ菇" & vbCrLf & vbCrLf & "�@�P�F�쐬���������܂�." & vbCrLf & vbCrLf & "�쐬���ʂ��\������܂��̂œ��e�ɏ]���Ă�������." & vbCrLf & vbCrLf & "�@�Q�F���M���������܂�." & vbCrLf & vbCrLf
Private mForm As New FormClass
Private mAbort As Boolean

Private Enum eCheckButton
    Yotei = 0
    Kakutei = 1
    Mukou = 2
End Enum

Private Sub cboFurikaeBi_Click()
    txtFurikaeBi.Text = cboFurikaeBi.Text
End Sub

Private Sub chkJisseki_Click()
    '//���т̎��͓��t�͕ύX�s�F�ŏI�̃f�[�^�ō쐬����
    txtFurikaeBi.Enabled = chkJisseki.Value = eCheckButton.Yotei
    cboFurikaeBi.Enabled = chkJisseki.Value = eCheckButton.Yotei
'//2004/04/13 �������ɂc�a�쐬��L���ɂ��違�e�L�X�g�쐬�E���M�𖳌��ɂ���F�c�a�쐬��L���ɁI
'//    cmdMakeDB.Enabled = chkJisseki.Value = eCheckButton.Yotei
    cmdMakeText.Enabled = chkJisseki.Value = eCheckButton.Yotei
    'cmdSend.Enabled = chkJisseki.Value = eCheckButton.Yotei

#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim MaxDay As Variant
    sql = "SELECT FASQNO,TO_CHAR(TO_DATE(FASQNO,'YYYYMMDD'),'YYYY/MM/DD') AS FaDate"
    sql = sql & " FROM tfFurikaeYoteiData"
    sql = sql & " GROUP BY FASQNO"
    sql = sql & " ORDER BY FASQNO"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    Call cboFurikaeBi.Clear
    Do Until dyn.EOF()
        Call cboFurikaeBi.AddItem(dyn.Fields("FaDate"))
        cboFurikaeBi.ItemData(cboFurikaeBi.NewIndex) = dyn.Fields("FASQNO")
        MaxDay = dyn.Fields("FASQNO")
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    '//�\��̎��͊�{���̎���U�֓���ǉ�
    If chkJisseki.Value = eCheckButton.Yotei Then
        sql = "SELECT AANXKZ,TO_CHAR(TO_DATE(AANXKZ,'YYYYMMDD'),'YYYY/MM/DD') AS AaDate"
        sql = sql & " FROM taSystemInformation"
        sql = sql & " WHERE AASKEY = 'SYSTEM'"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
        If Not dyn.EOF() Then
            '//�U�֗\��f�[�^�̍ŏI�����U�֓����傫�����̂�
            If MaxDay < dyn.Fields("AANXKZ") Then
                Call cboFurikaeBi.AddItem(dyn.Fields("AaDate"))
                cboFurikaeBi.ItemData(cboFurikaeBi.NewIndex) = dyn.Fields("AANXKZ")
            End If
        End If
    End If
    If cboFurikaeBi.ListCount Then
        cboFurikaeBi.ListIndex = cboFurikaeBi.ListCount - 1
    End If
    Dim ary As Variant
    ary = Array("(�\��)", "(����)")
    mCaption = Left(mCaption, IIf(InStr(mCaption, "("), InStr(mCaption, "(") - 1, Len(mCaption)))
    Me.Caption = Left(Me.Caption, IIf(InStr(Me.Caption, mCaption), InStr(Me.Caption, mCaption) - 1, Len(Me.Caption)))
    mCaption = mCaption & ary(chkJisseki.Value)
    Me.Caption = Me.Caption & mCaption
'//2004/04/13 �������ɂc�a�쐬��L���ɂ��違�e�L�X�g�쐬�E���M�𖳌��ɂ���F�c�a�쐬��L���ɁI
'//    cmdMakeText.Enabled = cboFurikaeBi.ListCount > 0
End Sub

Private Sub cmdEnd_Click()
    Unload Me
End Sub

#Const cSPEEDUP = True

Private Sub cmdMakeDB_Click()
    On Error GoTo cmdExport_ClickError
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim reg As New RegistryClass
    
'//2003/01/30 �ߋ��f�[�^���č쐬�ł��Ȃ�����
    If txtFurikaeBi.Text < gdDBS.sysDate("YYYY/MM/DD") Then
        Call MsgBox("�c�a�쐬�����悤�Ƃ��Ă�����t�͉ߋ��̓��t�ł�." & vbCrLf & vbCrLf & _
                    "�ߋ����t�f�[�^�͍쐬�ł��܂���." & vbCrLf & vbCrLf & _
                    "�T�[�o�[(" & reg.DbDatabaseName & ")���t = " & gdDBS.sysDate("YYYY/MM/DD"), vbInformation + vbOKOnly, mCaption)
        Exit Sub
    End If
'//2004/04/13 �������̗\��f�[�^�͍쐬�ł��Ȃ��悤�ɐ��䂷��B
'// If cboFurikaeBi.ListCount > 1 Then
    If cboFurikaeBi.ListIndex > 0 Then
        Call MsgBox("�������̂c�a�쐬(�\��)�͏o���܂���." & vbCrLf & vbCrLf & _
                    "��ɐU�֗\��\�̗ݐϏ��������s���Ă�������." _
                    , vbInformation + vbOKOnly, mCaption)
        Exit Sub
    End If
'// End If
    
    '//����_��҂�����������ƕی�҂����̌������̌��ʂ��Ԃ�̂� ==> DISTINCT
    sql = "SELECT DISTINCT a.ABITCD,c.* "
    sql = sql & " FROM taItakushaMaster     a,"
    sql = sql & "      tbKeiyakushaMaster   b,"
    '//��{�͕ی�҃}�X�^�[
    sql = sql & "      tcHogoshaMaster      c "
    sql = sql & " WHERE ABITKB = BAITKB"
    sql = sql & "   AND BAITKB = CAITKB"
    sql = sql & "   AND BAKYCD = CAKYCD"
'//2002/12/10 �����敪(??KSCD)�͎g�p���Ȃ�
'//    sql = sql & "   AND BAKSCD = CAKSCD"
    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN CAFKST AND CAFKED"
'//2003/02/03 ���t���O�Q�ƒǉ�
    sql = sql & "   AND NVL(BAKYFG,0) = 0"  '//�_��҂͉�񂵂Ă��Ȃ�
    sql = sql & "   AND NVL(CAKYFG,0) = 0"  '//�ی�҂͉�񂵂Ă��Ȃ�
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If dyn.EOF Then
        Call MsgBox(txtFurikaeBi.Text & " �ɊY������f�[�^�͂���܂���.", vbInformation + vbOKOnly, mCaption)
        Exit Sub
    End If
'//2003/02/03 �V�X�e���̃t���O���Q�Ƃ��Ă��悤�Ǝv�������������̃f�[�^���L��Əo���Ȃ��̂ł�߂�.
'//    If gdDBS.SystemUpdate("AAUPD2").Value <> 0 Then
'//        Call MsgBox(txtFurikaeBi.Text & " �ɊY������f�[�^�͂���܂���.", vbInformation + vbOKOnly, mCaption)
'//        Exit Sub
'//    End If
    
    Dim ms As New MouseClass
    Call ms.Start
    
'//2003/01/31 �V�K�G���g���[�f�[�^���f�p�V�X�e���L����
    Dim NewEntryStartDate As String, ReMake As Boolean
    NewEntryStartDate = Format(gdDBS.SystemUpdate("AANWDT"), "yyyy/mm/dd hh:nn:ss")
    
    Call gdDBS.Database.BeginTrans
    
    '//�֘A�e�[�u�����b�N�F2004/04/13 �{���Ƀ��b�N�ł���́H
    Call gdDBS.Database.ExecuteSQL("Lock Table tbKeiyakushaMaster IN EXCLUSIVE MODE NOWAIT")
    Call gdDBS.Database.ExecuteSQL("Lock Table tcHogoshaMaster    IN EXCLUSIVE MODE NOWAIT")
    Call gdDBS.Database.ExecuteSQL("Lock Table tfFurikaeYoteiData IN EXCLUSIVE MODE NOWAIT")
    Call gdDBS.Database.ExecuteSQL("Lock Table tfFurikaeYoteiTran IN EXCLUSIVE MODE NOWAIT")
    
    sql = "DELETE tfFurikaeYoteiData "
    sql = sql & " WHERE FASQNO = '" & txtFurikaeBi.Number & "'"
    If 0 <> gdDBS.Database.ExecuteSQL(sql) Then
        If vbYes <> MsgBox(txtFurikaeBi.Text & " �̃f�[�^�͊��ɑ��݂��܂�." & vbCrLf & vbCrLf & "�ēx�쐬���Ȃ��܂����H", vbInformation + vbDefaultButton3 + vbYesNoCancel, Me.Caption) Then
            GoTo cmdExport_ClickError
        End If
'//2003/02/03 �č쐬���͗\��쐬�����X�V���Ȃ�
        ReMake = True
    End If
    Dim cnt As Long

Debug.Print "start= " & Now

'////////////////////////////////////////////
'//2012/07/11 �X�s�[�h�A�b�v���P�F��������
#If cSPEEDUP = False Then
'''    Do Until dyn.EOF
'''        DoEvents
'''        If mAbort Then
'''            GoTo cmdExport_ClickError
'''        End If
'''        cnt = cnt + 1
'''        '//�U�֗\��f�[�^�ɒǉ�
'''        sql = "INSERT INTO tfFurikaeYoteiData VALUES("
''''//2003/01/31 Dynaset �� Object �Œ�`����� .Value ���t�����Ȃ��� Error=5 �ɂȂ�.
'''        sql = sql & "'" & dyn.Fields("CAITKB").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAKYCD").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAKSCD").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAHGCD").Value & "',"
'''        sql = sql & "'" & txtFurikaeBi.Number & "',"
'''        sql = sql & "'" & dyn.Fields("CAKKBN").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CABANK").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CASITN").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAKZSB").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAKZNO").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAYBTK").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAYBTN").Value & "',"
'''        sql = sql & "'" & dyn.Fields("CAKZNM").Value & "',"
'''        sql = sql & "'" & Val(gdDBS.Nz(dyn.Fields("CASKGK").Value)) & "',"
'''        sql = sql & "0,0,"                                  '//�ύX����z�E���t���O
''''//2003/01/31 �V�K�G���g���[�f�[�^���f�p�V�X�e���L��������ɔ��f
'''#If 0 Then
'''        '//�V�K�R�[�h 1=�V�K�ƂȂ�͂�...�B
'''        sql = sql & Abs(Format(NewEntryStartDate, "yyyy/mm/dd hh:nn:ss") _
'''                < Format(dyn.Fields("CAADDT").Value, "yyyy/mm/dd hh:nn:ss")) & ","
'''#Else
''''//2004/06/03 �c�a���ڂ�ǉ����Ĕ��f����悤�ɕύX�F�ݐώ��� CANWDT=SYSDATE �Ƃ��Ă���(�V�K����=NULL)
'''        sql = sql & Abs(IsNull(dyn.Fields("CANWDT").Value)) & ","
'''#End If
''''//2003/02/03 �X�V��ԃt���O�ǉ�
'''        sql = sql & eKouFuriKubun.YoteiDB & ","
'''        sql = sql & "'" & gdDBS.LoginUserName & "',"
'''        sql = sql & "SYSDATE"
'''        sql = sql & ")"
'''        Call gdDBS.Database.ExecuteSQL(sql)
'''        Call dyn.MoveNext
'''    Loop
#Else   'cSPEEDUP = False Then

    cnt = dyn.RecordCount
    
    sql = "INSERT INTO tfFurikaeYoteiData "
    sql = sql & "SELECT DISTINCT "
    sql = sql & "CAITKB,"
    sql = sql & "CAKYCD,"
    sql = sql & "CAKSCD,"
    sql = sql & "CAHGCD,"
    sql = sql & txtFurikaeBi.Number & ","
    sql = sql & "CAKKBN,"
    sql = sql & "CABANK,"
    sql = sql & "CASITN,"
    sql = sql & "CAKZSB,"
    sql = sql & "CAKZNO,"
    sql = sql & "CAYBTK,"
    sql = sql & "CAYBTN,"
    sql = sql & "CAKZNM,"
    sql = sql & "nvl(CASKGK,0),"
    sql = sql & "0,0,"
    sql = sql & "(case when CANWDT is null then 1 else 0 end),"
    sql = sql & eKouFuriKubun.YoteiDB & ","
    sql = sql & "'" & gdDBS.LoginUserName & "',"
    sql = sql & "SYSDATE"
    sql = sql & " FROM taItakushaMaster     a,"
    sql = sql & "      tbKeiyakushaMaster   b,"
    '//��{�͕ی�҃}�X�^�[
    sql = sql & "      tcHogoshaMaster      c "
    sql = sql & " WHERE ABITKB = BAITKB"
    sql = sql & "   AND BAITKB = CAITKB"
    sql = sql & "   AND BAKYCD = CAKYCD"
    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN CAFKST AND CAFKED"
    sql = sql & "   AND NVL(BAKYFG,0) = 0"  '//�_��҂͉�񂵂Ă��Ȃ�
    sql = sql & "   AND NVL(CAKYFG,0) = 0"  '//�ی�҂͉�񂵂Ă��Ȃ�
    Dim insCnt As Long
    insCnt = gdDBS.Database.ExecuteSQL(sql)
    If insCnt <> cnt Then
        Call Err.Raise(-1, "cmdMakeDB", "�c�a�쐬�͎��s���܂���.")
    End If
#End If     'cSPEEDUP = False Then
'//2012/07/11 �X�s�[�h�A�b�v���P�F�����܂�
'////////////////////////////////////////////

Debug.Print "  end= " & Now
    
    Call dyn.Close
    '//���s�X�V�t���O�ݒ�F���̊֐��͗\��̂ݎ��s�\
'//2003/02/03 �č쐬���͗\��쐬�����X�V���Ȃ�
    If ReMake = False Then
        gdDBS.SystemUpdate("AAUPD1") = 1
    End If
    
'//2004/05/17 �ڍׂ��֐���
    Call pNormalEndMessage(ReMake, cnt, NewEntryStartDate)
    
    Call gdDBS.Database.CommitTrans
    
'//2004/04/13 �������ɂc�a�쐬��L���ɂ��違�e�L�X�g�쐬�E���M�𖳌��ɂ���F�c�a�쐬��L���ɁI
    cmdMakeText.Enabled = True
    'cmdSend.Enabled = True
    Exit Sub
cmdExport_ClickError:
    Call gdDBS.Database.Rollback
    Call gdDBS.ErrorCheck(gdDBS.Database)       '//�G���[�g���b�v
'// gdDBS.ErrorCheck() �̏�Ɉړ�
'//    Call gdDBS.Database.Rollback
End Sub

Private Sub pNormalEndMessage(ByVal vRemake As Boolean, vCnt As Long, ByVal vNewEntryStartDate As Variant, Optional vMsgMode As Boolean = False)
'//2004/05/17 �ڍׂ��֐���
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If

'//2007/07/18 �����\����S�ʓI�Ɍ�����
    Dim allCnt As Long, ttlGaku As Currency
    Dim oldNew As Long, oldZero As Long, oldCan As Long, canCnt As Long
    Dim newNew As Long, newZero As Long, newCan As Long, ssCancel As Long
    
    sql = "SELECT " & vbCrLf
    sql = sql & " SUM(oldNew)   oldNew  ," & vbCrLf     '//�ߋ��̐V�K����(�V�K������)
    sql = sql & " SUM(oldZero)  oldZero ," & vbCrLf     '//�ߋ��̂O�~����(�V�K������)
    sql = sql & " SUM(oldCan)   oldCan  ," & vbCrLf     '//�ߋ��̍����񌏐�(�V�K������)
    sql = sql & " SUM(canCnt)   canCnt  ," & vbCrLf     '//�ߋ���    ��񌏐�(�V�K������)
    
    sql = sql & " SUM(newNew)   newNew  ," & vbCrLf     '//����̐V�K����
    sql = sql & " SUM(newZero)  newZero ," & vbCrLf     '//����̂O�~����
    sql = sql & " SUM(newCan)   newCan  ," & vbCrLf     '//����̉�񌏐�
    
    sql = sql & " SUM(allCnt)   allCnt  ," & vbCrLf     '//�����f�[�^�̑�����
    sql = sql & " SUM(TtlGaku)  TtlGaku ," & vbCrLf     '//���������z
    sql = sql & " SUM(ssCancel) ssCancel " & vbCrLf     '//�搶�F�_��҂̃L�����Z���ŕی�Ґ����f�[�^
    
    sql = sql & " FROM (" & vbCrLf
    '////////////////////////////////////
    '//�U�֗\��f�[�^���擾������e�F�������̃f�[�^
    sql = sql & " SELECT " & vbCrLf
    sql = sql & " 0                                     oldNew  ," & vbCrLf     '//�ߋ��̐V�K����(�V�K������)
    sql = sql & " 0                                     oldZero ," & vbCrLf     '//�ߋ��̂O�~����(�V�K������)
    sql = sql & " 0                                     oldCan  ," & vbCrLf     '//�ߋ��̍����񌏐�(�V�K������)
    sql = sql & " 0                                     canCnt  ," & vbCrLf     '//�ߋ���    ��񌏐�(�V�K������)
    sql = sql & " 0                                     newNew  ," & vbCrLf     '//����̐V�K����
    sql = sql & " 0                                     newZero ," & vbCrLf     '//����̂O�~����
    sql = sql & " 0                                     newCan  ," & vbCrLf     '//����̉�񌏐�
    sql = sql & " COUNT(*)                              allCnt  ," & vbCrLf     '//�����f�[�^�̑�����
    sql = sql & " SUM(NVL(FASKGK,0))                    TtlGaku ," & vbCrLf     '//���������z
    sql = sql & " 0                                     ssCancel " & vbCrLf     '//�搶�F�_��҂̃L�����Z���ŕی�Ґ����f�[�^
    sql = sql & " FROM tfFurikaeYoteiData " & vbCrLf
    sql = sql & " WHERE FASQNO = '" & txtFurikaeBi.Number & "'" & vbCrLf
    sql = sql & " UNION ALL " & vbCrLf
    '////////////////////////////////////
    '//�U�֗\��f�[�^���擾������e�F�ߋ��̐V�K�f�[�^
    sql = sql & " SELECT " & vbCrLf
    sql = sql & " SUM("
    sql = sql & "   CASE WHEN NVL(FASKGK,0) <> 0 THEN "
    sql = sql & "         DECODE(NVL(FANWCD,0),0,0,1) "
    sql = sql & "   END"
    sql = sql & " )                                     oldNew  ," & vbCrLf     '//�ߋ��̐V�K����(�V�K������)
    sql = sql & " SUM("
    sql = sql & "   CASE WHEN NVL(FASKGK,0)  = 0 THEN "
    sql = sql & "         DECODE(NVL(FANWCD,0),0,0,1) "
    sql = sql & "   END"
    sql = sql & " )                                     oldZero ," & vbCrLf     '//�ߋ��̂O�~����(�V�K������)
    sql = sql & " 0                                     oldCan  ," & vbCrLf     '//�ߋ��̍����񌏐�(�V�K������)
    sql = sql & " 0                                     canCnt  ," & vbCrLf     '//�ߋ���    ��񌏐�(�V�K������)
    sql = sql & " 0                                     newNew  ," & vbCrLf     '//����̐V�K����
    sql = sql & " 0                                     newZero ," & vbCrLf     '//����̂O�~����
    sql = sql & " 0                                     newCan  ," & vbCrLf     '//����̉�񌏐�
    sql = sql & " 0                                     allCnt  ," & vbCrLf     '//�����f�[�^�̑�����
    sql = sql & " 0                                     TtlGaku ," & vbCrLf     '//���������z
    sql = sql & " 0                                     ssCancel " & vbCrLf     '//�搶�F�_��҂̃L�����Z���ŕی�Ґ����f�[�^
    sql = sql & " FROM tfFurikaeYoteiData " & vbCrLf
    sql = sql & " WHERE FASQNO = '" & txtFurikaeBi.Number & "'" & vbCrLf
    sql = sql & "   AND       (FAITKB,FAKYCD,FAKSCD,FAHGCD) IN (" & vbCrLf
    sql = sql & "       SELECT CAITKB,CAKYCD,CAKSCD,CAHGCD " & vbCrLf
    sql = sql & "       FROM tcHogoshaMaster    a," & vbCrLf
    sql = sql & "            tbKeiyakushaMaster b " & vbCrLf
    sql = sql & "       WHERE CAITKB = BAITKB " & vbCrLf
    sql = sql & "         AND CAKYCD = BAKYCD " & vbCrLf
    sql = sql & "         AND NVL(BAKYFG,0) = 0 " & vbCrLf  '//�_��҂�����ԂłȂ��I
    sql = sql & "         AND " & txtFurikaeBi.Number & " BETWEEN BAKYST AND BAKYED " & vbCrLf
    sql = sql & "         AND " & txtFurikaeBi.Number & " BETWEEN BAFKST AND BAFKED " & vbCrLf
    sql = sql & "         AND CAADDT < TO_DATE('" & vNewEntryStartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
    sql = sql & "         AND CANWDT IS NULL " & vbCrLf
'    sql = sql & "         AND " & txtFurikaeBi.Number & " BETWEEN CAKYST AND CAKYED " & vbCrLf
'    sql = sql & "         AND " & txtFurikaeBi.Number & " BETWEEN CAFKST AND CAFKED " & vbCrLf
    sql = sql & "       )" & vbCrLf
    sql = sql & " UNION ALL " & vbCrLf
    '////////////////////////////////////
    '//���͕ی�҃}�X�^���擾����F�ߋ��̉��f�[�^
    sql = sql & " SELECT " & vbCrLf
    sql = sql & " 0                                     oldNew  ," & vbCrLf     '//�ߋ��̐V�K����(�V�K������)
    sql = sql & " 0                                     oldZero ," & vbCrLf     '//�ߋ��̂O�~����(�V�K������)
    sql = sql & " SUM(" & vbCrLf
    sql = sql & "   CASE WHEN CAKYSR >= TO_DATE('" & vNewEntryStartDate & "','YYYY/MM/DD HH24:MI:SS') THEN 1 " & vbCrLf
    sql = sql & "   ELSE 0 " & vbCrLf
    sql = sql & "   END" & vbCrLf
    sql = sql & " ) AS                                  oldCan  ," & vbCrLf     '//�ߋ��̍����񌏐�(�V�K������)
    
    sql = sql & " SUM(" & vbCrLf
    sql = sql & "   CASE WHEN CAKYSR IS NULL THEN 1 " & vbCrLf
    sql = sql & "        WHEN CAKYSR < TO_DATE('" & vNewEntryStartDate & "','YYYY/MM/DD HH24:MI:SS') THEN 1 " & vbCrLf
    sql = sql & "   ELSE 0 " & vbCrLf
    sql = sql & "   END" & vbCrLf
    sql = sql & " ) AS                                  canCnt  ," & vbCrLf     '//�ߋ���    ��񌏐�(�V�K������)
    sql = sql & " 0                                     newNew  ," & vbCrLf     '//����̐V�K����
    sql = sql & " 0                                     newZero ," & vbCrLf     '//����̂O�~����
    sql = sql & " 0                                     newCan  ," & vbCrLf     '//����̉�񌏐�
    sql = sql & " 0                                     allCnt  ," & vbCrLf     '//�����f�[�^�̑�����
    sql = sql & " 0                                     TtlGaku ," & vbCrLf     '//���������z
    sql = sql & " 0                                     ssCancel " & vbCrLf     '//�搶�F�_��҂̃L�����Z���ŕی�Ґ����f�[�^
    sql = sql & " FROM tcHogoshaMaster    a," & vbCrLf
    sql = sql & "      tbKeiyakushaMaster b " & vbCrLf
    sql = sql & " WHERE CAITKB = BAITKB " & vbCrLf
    sql = sql & "   AND CAKYCD = BAKYCD " & vbCrLf
    sql = sql & "   AND NVL(BAKYFG,0) = 0 " & vbCrLf    '//�_��҂�����ԂłȂ��I
    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN BAKYST AND BAKYED " & vbCrLf
    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN BAFKST AND BAFKED " & vbCrLf
    sql = sql & "   AND CAADDT < TO_DATE('" & vNewEntryStartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
    sql = sql & "   AND CANWDT IS NULL " & vbCrLf
    sql = sql & "   AND NVL(CAKYFG,0) <> 0 " & vbCrLf   '//�ی�҂͉���ԁI
'    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN CAKYST AND CAKYED " & vbCrLf
'    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN CAFKST AND CAFKED " & vbCrLf
    sql = sql & " UNION ALL " & vbCrLf
    '////////////////////////////////////
    '//�U�֗\��f�[�^���擾������e�F����̐V�K�f�[�^
    sql = sql & " SELECT " & vbCrLf
    sql = sql & " 0                                     oldNew  ," & vbCrLf     '//�ߋ��̐V�K����(�V�K������)
    sql = sql & " 0                                     oldZero ," & vbCrLf     '//�ߋ��̂O�~����(�V�K������)
    sql = sql & " 0                                     oldCan  ," & vbCrLf     '//�ߋ��̍����񌏐�(�V�K������)
    sql = sql & " 0                                     canCnt  ," & vbCrLf     '//�ߋ���    ��񌏐�(�V�K������)
    sql = sql & " SUM("
    sql = sql & "   CASE WHEN NVL(FASKGK,0) <> 0 THEN "
    sql = sql & "         DECODE(NVL(FANWCD,0),0,0,1) "
    sql = sql & "   END"
    sql = sql & " )                                     newNew  ," & vbCrLf     '//����̐V�K����
    sql = sql & " SUM("
    sql = sql & "   CASE WHEN NVL(FASKGK,0)  = 0 THEN "
    sql = sql & "         DECODE(NVL(FANWCD,0),0,0,1) "
    sql = sql & "   END"
    sql = sql & " )                                     newZero ," & vbCrLf     '//����̂O�~����
    sql = sql & " 0                                     newCan  ," & vbCrLf     '//����̉�񌏐�
    sql = sql & " 0                                     allCnt  ," & vbCrLf     '//�����f�[�^�̑�����
    sql = sql & " 0                                     TtlGaku ," & vbCrLf     '//���������z
    sql = sql & " 0                                     ssCancel " & vbCrLf     '//�搶�F�_��҂̃L�����Z���ŕی�Ґ����f�[�^
    sql = sql & " FROM tfFurikaeYoteiData " & vbCrLf
    sql = sql & " WHERE FASQNO = '" & txtFurikaeBi.Number & "'" & vbCrLf
    sql = sql & "   AND       (FAITKB,FAKYCD,FAKSCD,FAHGCD) IN (" & vbCrLf
    sql = sql & "       SELECT CAITKB,CAKYCD,CAKSCD,CAHGCD " & vbCrLf
    sql = sql & "       FROM tcHogoshaMaster    a," & vbCrLf
    sql = sql & "            tbKeiyakushaMaster b " & vbCrLf
    sql = sql & "       WHERE CAITKB = BAITKB " & vbCrLf
    sql = sql & "         AND CAKYCD = BAKYCD " & vbCrLf
    sql = sql & "         AND NVL(BAKYFG,0) = 0 " & vbCrLf  '//�_��҂�����ԂłȂ��I
    sql = sql & "         AND " & txtFurikaeBi.Number & " BETWEEN BAKYST AND BAKYED " & vbCrLf
    sql = sql & "         AND " & txtFurikaeBi.Number & " BETWEEN BAFKST AND BAFKED " & vbCrLf
    sql = sql & "         AND CAADDT >= TO_DATE('" & vNewEntryStartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
    sql = sql & "         AND CANWDT IS NULL " & vbCrLf
'    sql = sql & "         AND " & txtFurikaeBi.Number & " BETWEEN CAKYST AND CAKYED " & vbCrLf
'    sql = sql & "         AND " & txtFurikaeBi.Number & " BETWEEN CAFKST AND CAFKED " & vbCrLf
    sql = sql & "       )" & vbCrLf
    sql = sql & " UNION ALL " & vbCrLf
    '////////////////////////////////////
    '//���͕ی�҃}�X�^���擾����F����̉��f�[�^
    sql = sql & " SELECT " & vbCrLf
    sql = sql & " 0                                     oldNew  ," & vbCrLf     '//�ߋ��̐V�K����(�V�K������)
    sql = sql & " 0                                     oldZero ," & vbCrLf     '//�ߋ��̂O�~����(�V�K������)
    sql = sql & " 0                                     oldCan  ," & vbCrLf     '//�ߋ��̍����񌏐�(�V�K������)
    sql = sql & " 0                                     canCnt  ," & vbCrLf     '//�ߋ���    ��񌏐�(�V�K������)
    sql = sql & " 0                                     newNew  ," & vbCrLf     '//����̐V�K����
    sql = sql & " 0                                     newZero ," & vbCrLf     '//����̂O�~����
    sql = sql & " SUM(DECODE(NVL(CAKYFG,0),0,0,1))      newCan  ," & vbCrLf     '//����̉�񌏐�
    sql = sql & " 0                                     allCnt  ," & vbCrLf     '//�����f�[�^�̑�����
    sql = sql & " 0                                     TtlGaku ," & vbCrLf     '//���������z
    sql = sql & " 0                                     ssCancel " & vbCrLf     '//�搶�F�_��҂̃L�����Z���ŕی�Ґ����f�[�^
    sql = sql & " FROM tcHogoshaMaster    a," & vbCrLf
    sql = sql & "      tbKeiyakushaMaster b " & vbCrLf
    sql = sql & " WHERE CAITKB = BAITKB " & vbCrLf
    sql = sql & "   AND CAKYCD = BAKYCD " & vbCrLf
    sql = sql & "   AND NVL(BAKYFG,0) = 0 " & vbCrLf  '//�_��҂�����ԂłȂ��I
    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN BAKYST AND BAKYED " & vbCrLf
    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN BAFKST AND BAFKED " & vbCrLf
    sql = sql & "   AND CAADDT >= TO_DATE('" & vNewEntryStartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
    sql = sql & "   AND CANWDT IS NULL " & vbCrLf
    sql = sql & "   AND NVL(CAKYFG,0) <> 0 " & vbCrLf   '//�ی�҂͉���ԁI
'    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN CAKYST AND CAKYED " & vbCrLf
'    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN CAFKST AND CAFKED " & vbCrLf
    sql = sql & " UNION ALL " & vbCrLf
    '////////////////////////////////////
    '//�搶�F�_��҂̃L�����Z���ŕی�Ґ����f�[�^
    sql = sql & " SELECT " & vbCrLf
    sql = sql & " 0                                     oldNew  ," & vbCrLf     '//�ߋ��̐V�K����(�V�K������)
    sql = sql & " 0                                     oldZero ," & vbCrLf     '//�ߋ��̂O�~����(�V�K������)
    sql = sql & " 0                                     oldCan  ," & vbCrLf     '//�ߋ��̍����񌏐�(�V�K������)
    sql = sql & " 0                                     canCnt  ," & vbCrLf     '//�ߋ���    ��񌏐�(�V�K������)
    sql = sql & " 0                                     newNew  ," & vbCrLf     '//����̐V�K����
    sql = sql & " 0                                     newZero ," & vbCrLf     '//����̂O�~����
    sql = sql & " 0                                     newCan  ," & vbCrLf     '//����̉�񌏐�
    sql = sql & " 0                                     allCnt  ," & vbCrLf     '//�����f�[�^�̑�����
    sql = sql & " 0                                     TtlGaku ," & vbCrLf     '//���������z
    sql = sql & " COUNT(*)                              ssCancel " & vbCrLf     '//�搶�F�_��҂̃L�����Z���ŕی�Ґ����f�[�^
    sql = sql & " FROM tcHogoshaMaster    a," & vbCrLf
    sql = sql & "      tbKeiyakushaMaster b " & vbCrLf
    sql = sql & " WHERE CAITKB = BAITKB " & vbCrLf
    sql = sql & "   AND CAKYCD = BAKYCD " & vbCrLf
    sql = sql & "   AND NVL(BAKYFG,0) <> 0 " & vbCrLf  '//�_��҂�����ԁI
'    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN BAKYST AND BAKYED " & vbCrLf
'    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN BAFKST AND BAFKED " & vbCrLf
'    sql = sql & "   AND CAADDT >= TO_DATE('" & vNewEntryStartDate & "','YYYY/MM/DD HH24:MI:SS')" & vbCrLf
    sql = sql & "   AND CANWDT IS NULL " & vbCrLf
    sql = sql & "   AND NVL(CAKYFG,0) = 0 " & vbCrLf   '//�ی�҂͉���ԂłȂ��I
'    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN CAKYST AND CAKYED " & vbCrLf
'    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN CAFKST AND CAFKED " & vbCrLf
    sql = sql & ")"
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
    If Not dyn.EOF Then
        oldNew = dyn.Fields("oldNew").Value
        oldZero = dyn.Fields("oldZero").Value
        oldCan = dyn.Fields("oldCan").Value
        canCnt = dyn.Fields("canCnt").Value
        newNew = dyn.Fields("newNew").Value
        newZero = dyn.Fields("newZero").Value
        newCan = dyn.Fields("newCan").Value
        allCnt = dyn.Fields("allCnt").Value
        ssCancel = dyn.Fields("ssCancel").Value
        ttlGaku = dyn.Fields("ttlGaku").Value
    End If
    Call dyn.Close
    
'//2004/04/26 �V�K����&�O�~���������v���z�̒ǉ�
'//2004/05/17 ���O�~�J�E���g��V�K�̂O�~�ɕύX
'//2006/04/25 �V�K��񌏐��J�E���g�ǉ�
'//2007/03/08 ���b�Z�[�W��\������{�^���ǉ��̂��߃��O�o�͂��Ȃ�
'//2007/07/18 �����\����S�ʓI�Ɍ�����
    If False = vMsgMode Then
        Call gdDBS.AutoLogOut(mCaption, "�c�a" & IIf(vRemake = True, "��", "�V�K") & "�쐬(" & _
                    "�����U�֓�=[" & txtFurikaeBi.Text & "] �쐬������=" & vCnt & " �V�K�����̏ڍ� ==> " & _
                    " �O��ȑO = <����=" & oldNew & " : �O�~=" & oldZero & " : ���=" & oldCan & ">" & _
                    " ����ǉ� = <����=" & newNew & " : �O�~=" & newZero & " : ���=" & newCan & ">" & _
                    " �_��҉��ŕی�ҐV�K�����f�[�^=" & ssCancel)
    End If
'//2004/04/26 �V�K����&�O�~���������v���z�̒ǉ�
'//2004/05/17 ���O�~�J�E���g��V�K�̂O�~�ɕύX
'//2006/04/25 �V�K��񌏐��J�E���g�ǉ�
'//2007/07/18 �����\����S�ʓI�Ɍ�����
    Dim st As New StringClass
    lblMessage.Caption = Format(vCnt, "#,0") & " ���̃f�[�^���쐬����܂����B" & vbCrLf & vbCrLf & _
                        "<< �V�K�����̏ڍ� >>" & vbCrLf & _
                        Space(3) & Space(16) & "����" & Space(6) & "�O�~" & Space(2) & "������" & Space(1) & "(�ߋ����)" & Space(5) & "���v" & vbCrLf & _
                        Space(3) & String(60, "=") & vbCrLf & _
                        Space(3) & "�O��ȑO =" & st.FixedFormat(oldNew, 10) & st.FixedFormat(oldZero, 10) & st.FixedFormat(oldCan, 10) & st.FixedFormat(canCnt, 10) & st.FixedFormat(oldNew + oldZero + oldCan + canCnt, 10) & vbCrLf & _
                        Space(3) & String(60, "-") & vbCrLf & _
                        Space(3) & "����ǉ� =" & st.FixedFormat(newNew, 10) & st.FixedFormat(newZero, 10) & st.FixedFormat(newCan, 10) & Space(10) & st.FixedFormat(newNew + newZero + newCan, 10) & vbCrLf & _
                        Space(3) & String(60, "=") & vbCrLf & _
                        Space(3) & "�V�K���v =" & st.FixedFormat(oldNew + newNew, 10) & st.FixedFormat(oldZero + newZero, 10) & st.FixedFormat(oldCan + newCan, 10) & st.FixedFormat(canCnt, 10) & vbCrLf & _
                        Space(3) & String(60, "=") & vbCrLf & _
                        Space(5) & "�쐬���ꂽ�������� " & Format(allCnt, "#,0") & " ���ł��B"
    If 0 <> ssCancel Then
        lblMessage.Caption = lblMessage.Caption & vbCrLf & vbCrLf & " �� �_��҉��ŕی�ҐV�K�����f�[�^�� " & ssCancel & " �����݂��܂��B"
    End If

'//2007/03/08 ���b�Z�[�W��\������{�^���ǉ��̂��߃��O�o�͂��Ȃ�
    If True = vMsgMode Then
        lblMessage.Caption = "====== �O��̍쐬���� =====" & vbCrLf & lblMessage.Caption
    Else
        Call MsgBox(IIf(vRemake = True, "��", "�V�K") & "�쐬�͐���I�����܂����B" & vbCrLf & vbCrLf & "�o�̓��b�Z�[�W�̓��e���m�F���ĉ������B", vbInformation, mCaption)
    End If
    lblMessage.AutoSize = True
End Sub

Private Sub cmdMakeText_Click()
    On Error GoTo cmdExport_ClickError
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    
    sql = "SELECT a.ABITCD,c.*,c.rowid "
    sql = sql & " FROM taItakushaMaster     a,"
    sql = sql & "      tcHogoshaMaster      b,"
    '//��{�͐U�֗\��f�[�^
    sql = sql & "      tfFurikaeYoteiData   c "
    sql = sql & " WHERE ABITKB = FAITKB"
    sql = sql & "   AND FAITKB = CAITKB"
    sql = sql & "   AND FAKYCD = CAKYCD"
    sql = sql & "   AND FAKSCD = CAKSCD"
    sql = sql & "   AND FAHGCD = CAHGCD"
    sql = sql & "   AND " & txtFurikaeBi.Number & " BETWEEN CAFKST AND CAFKED"
    sql = sql & "   AND FASQNO = " & txtFurikaeBi.Number
'//2003/02/03 ���t���O�Q�ƒǉ�
    sql = sql & "   AND NVL(FAKYFG,0) = 0"  '//�ی�҂͉�񂵂Ă��Ȃ�
'//2004/06/03 ���z�u�O�v�͍쐬���Ȃ�
'//2004/06/03 �^�p���ς��H�̂Ŏ~�߁I�I�I
'    sql = sql & "   AND(NVL(faskgk,0) > 0 OR NVL(fahkgk,0) > 0) "
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If dyn.EOF Then
        Call MsgBox(txtFurikaeBi.Text & " �ɊY������f�[�^�͂���܂���.", vbInformation + vbOKOnly, mCaption)
        Exit Sub
    End If
    Dim st As New StructureClass, tmp As String
    Dim reg As New RegistryClass
    Dim mFile As New FileClass, FileName As String, TmpFname As String
    
    dlgFile.DialogTitle = "���O��t���ĕۑ�(" & mCaption & ")"
    dlgFile.FileName = reg.OutputFileName(mCaption)
    If IsEmpty(mFile.SaveDialog(dlgFile)) Then
        Exit Sub
    End If
    
    Dim ms As New MouseClass
    Call ms.Start
    
    reg.OutputFileName(mCaption) = dlgFile.FileName
    Call st.SelectStructure(st.KouzaFurikae)

Debug.Print "start= " & Now
    Me.pgbRecord.Visible = True
    Me.pgbRecord.ZOrder 0
    Me.pgbRecord.min = 0
    Me.pgbRecord.Max = dyn.RecordCount
    
    '//��芸�����e���|�����ɏ���
    Dim fp As Integer, cnt As Long, SumGaku As Currency
    fp = FreeFile
    TmpFname = mFile.MakeTempFile
    Open TmpFname For Append As #fp
    Do Until dyn.EOF
        DoEvents
        If mAbort Then
            GoTo cmdExport_ClickError
        End If
        tmp = ""
        tmp = tmp & st.SetData(dyn.Fields("ABITCD"), 0)     '�ϑ��Ҕԍ�             '//���̍��ڂ͈ϑ��҃}�X�^
        tmp = tmp & st.SetData(dyn.Fields("FAKYCD"), 1)     '�_��Ҕԍ�(����)
        tmp = tmp & st.SetData(dyn.Fields("FAKSCD"), 2)     '�����敪
        tmp = tmp & st.SetData("000", 3)                    '�[���X�y�[�X�Œ�F�����敪�ł͂Ȃ�
        tmp = tmp & st.SetData(dyn.Fields("FAHGCD"), 4)     '�ی�Ҕԍ�
        '//2002/11/26 �󔒂T�����ǉ�
        tmp = tmp & String(5, " ")
        '//���Z�@�ւ̋敪�ɂ���ċ�s���X�֋ǂ̌��ʂ�ԋp����֐��� StructureClass ���쐬
        tmp = tmp & st.SetData(st.BankCode(dyn), 5)         '��s�R�[�h
        tmp = tmp & st.SetData(st.ShitenCode(dyn), 6)       '�x�X�R�[�h
        tmp = tmp & st.SetData(st.Shubetsu(dyn), 7)         '�a�����
        tmp = tmp & st.SetData(st.KouzaNo(dyn), 8)          '�����ԍ�
        '//���Z�@�ւ̋敪�ɂ���ċ�s���X�֋ǂ̌��ʂ�ԋp����֐��� StructureClass ���쐬
        tmp = tmp & st.SetData(dyn.Fields("FAKZNM"), 9)     '�������`�l��(�J�i)
        tmp = tmp & st.SetData(dyn.Fields("FASKGK"), 10)    '�������z
        SumGaku = SumGaku + Val(gdDBS.Nz(dyn.Fields("FASKGK")))
'//���������ĐV�K�E���̑������߂�H
'//�V�K�R�[�h 1=�V�K
        tmp = tmp & st.SetData(Val(gdDBS.Nz(dyn.Fields("FANWCD"))), 11)  '�V�K�R�[�h �V�K="1",���̑�="0"
        Print #fp, tmp
        cnt = cnt + 1
        Me.pgbRecord.Value = cnt
'////////////////////////////////////////////
'//2012/07/11 �X�s�[�h�A�b�v���P�F��������
#If cSPEEDUP = False Then
''''//2003/02/03 �X�V��ԃt���O�ǉ�:0=DB�쐬,1=�\��쐬,2=�\��捞,3=�����쐬
'''        sql = "UPDATE tfFurikaeYoteiData SET "
'''        sql = sql & " FAUPFG = " & IIf(chkJisseki.Value = eCheckButton.Yotei, _
'''                                        eKouFuriKubun.YoteiText, _
'''                                        eKouFuriKubun.SeikyuText _
'''                                ) & ","
'''        sql = sql & " FAUSID = '" & gdDBS.LoginUserName & "',"
'''        sql = sql & " FAUPDT = SYSDATE"
'''        sql = sql & " WHERE FAITKB = '" & dyn.Fields("FAITKB").Value & "'"
'''        sql = sql & "   AND FAKYCD = '" & dyn.Fields("FAKYCD").Value & "'"
'''        sql = sql & "   AND FAKSCD = '" & dyn.Fields("FAKSCD").Value & "'"
'''        sql = sql & "   AND FAHGCD = '" & dyn.Fields("FAHGCD").Value & "'"
'''        sql = sql & "   AND FASQNO = " & txtFurikaeBi.Number
'''        Call gdDBS.Database.ExecuteSQL(sql)
#End If
'//2012/07/11 �X�s�[�h�A�b�v���P�F�����܂�
'////////////////////////////////////////////
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    Me.pgbRecord.Visible = False
    lblMessage.ZOrder 0
    Me.Refresh
'////////////////////////////////////////////
'//2012/07/11 �X�s�[�h�A�b�v���P�F��������
#If cSPEEDUP = True Then
    sql = "UPDATE tfFurikaeYoteiData SET "
    sql = sql & " FAUPFG = " & IIf(chkJisseki.Value = eCheckButton.Yotei, _
                                    eKouFuriKubun.YoteiText, _
                                    eKouFuriKubun.SeikyuText _
                            ) & ","
    sql = sql & " FAUSID = '" & gdDBS.LoginUserName & "',"
    sql = sql & " FAUPDT = SYSDATE"
    sql = sql & " WHERE FASQNO = " & txtFurikaeBi.Number
    Dim updCnt As Long
    updCnt = gdDBS.Database.ExecuteSQL(sql)
    If updCnt <> cnt Then
        Call Err.Raise(-1, "cmdMakeDB", "�e�L�X�g�쐬�͎��s���܂���.")
    End If
#End If
'//2012/07/11 �X�s�[�h�A�b�v���P�F�����܂�
'////////////////////////////////////////////

Debug.Print "  end= " & Now

#If 0 Then
'//2004/04/26 �V�K����&�O�~���������v���z�̒ǉ�
'//2004/05/17 �ڍׂ��폜
    Dim newCnt As Long, ZeroCnt As Long, TotalGaku As Currency
    sql = "SELECT "
    sql = sql & " SUM(NVL(FANWCD,0)) AS NewCnt,"
    sql = sql & " SUM(DECODE(FASKGK,0,1,0)) AS ZeroCnt,"
    sql = sql & " SUM(NVL(FASKGK,0)) AS TotalGaku "
    sql = sql & " FROM tfFurikaeYoteiData"
    sql = sql & " WHERE FASQNO = '" & txtFurikaeBi.Number & "'"
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
    If Not dyn.EOF Then
        newCnt = dyn.Fields("NewCnt").Value
        ZeroCnt = dyn.Fields("ZeroCnt").Value
        TotalGaku = dyn.Fields("TotalGaku").Value
    End If
    Call dyn.Close
#End If
#If NO_TOTAL_REC Then
'//2003/02/03 ��������@���v�����E���z���R�[�h�ǉ�
    tmp = ""
    tmp = tmp & st.SetData("9999999999", 0)     '�ϑ��Ҕԍ�             '//���̍��ڂ͈ϑ��҃}�X�^
    tmp = tmp & st.SetData("9999999999", 1)     '�_��Ҕԍ�(����)
    tmp = tmp & st.SetData("9999999999", 2)     '�����敪
    tmp = tmp & st.SetData("9999999999", 3)                    '�[���X�y�[�X�Œ�F�����敪�ł͂Ȃ�
    tmp = tmp & st.SetData("9999999999", 4)     '�ی�Ҕԍ�
    tmp = tmp & String(5, " ")                  '�󔒂T�����ǉ�
    '//���Z�@�ւ̋敪�ɂ���ċ�s���X�֋ǂ̌��ʂ�ԋp����֐��� StructureClass ���쐬
    tmp = tmp & st.SetData("", 5)     '��s�R�[�h
    tmp = tmp & st.SetData("", 6)     '�x�X�R�[�h
    tmp = tmp & st.SetData("", 7)     '�a�����
    tmp = tmp & st.SetData(cnt, 8)     '������ ���v���� ������ �����ԍ�
    '//���Z�@�ւ̋敪�ɂ���ċ�s���X�֋ǂ̌��ʂ�ԋp����֐��� StructureClass ���쐬
    tmp = tmp & st.SetData("�޳��(�ݽ�/�ݶ޸)ں���", 9)     '�������`�l��(�J�i)
    tmp = tmp & st.SetData(SumGaku, 10)         '������ ���v���z ������ �������z
    tmp = tmp & st.SetData("0", 11)             '�V�K�R�[�h �V�K="1",���̑�="0"
    Print #fp, tmp
'//2003/02/03 �����܂Ł@���v�����E���z���R�[�h�ǉ�
#End If
    Close #fp
#If 1 Then
    '//�t�@�C���ړ�     MOVEFILE_REPLACE_EXISTING=Replace , MOVEFILE_COPY_ALLOWED=Copy & Delete
    Call MoveFileEx(TmpFname, reg.OutputFileName(mCaption), MOVEFILE_REPLACE_EXISTING + MOVEFILE_COPY_ALLOWED)
    'Call MoveFileEx(TmpFname, reg.FileName(mCaption), MOVEFILE_REPLACE_EXISTING)
#Else
    '//�t�@�C���R�s�[
    Call FileCopy(TmpFname, reg.FileName(mCaption))
#End If
    Set mFile = Nothing
    '//���s�X�V�t���O�ݒ�F���̊֐��͗\��E�����Ƃ��Ɏ��s�\
    Select Case chkJisseki.Value
    Case eCheckButton.Yotei
        gdDBS.SystemUpdate("AAUPD2") = 1
    Case eCheckButton.Kakutei
        gdDBS.SystemUpdate("AAUPD3") = 1
    End Select
    Call gdDBS.AutoLogOut(mCaption, "�e�L�X�g�쐬(" & txtFurikaeBi & " : " & cnt & " ��)")
'//2004/04/26 �V�K����&�O�~���������v���z�̒ǉ�
'//2004/05/17 �ڍׂ��폜
    lblMessage.Caption = cnt & " ���̃f�[�^���쐬����܂����B"
                    '// & vbCrLf & _
                        "<< �ڍ� >>" & vbCrLf & _
                        "�V�K���� = " & NewCnt & vbCrLf & _
                        "  �O�~���� = " & ZeroCnt & vbCrLf & _
                        "���v���z = " & Format(TotalGaku, "#,##0")
    Exit Sub
cmdExport_ClickError:
    Call gdDBS.ErrorCheck       '//�G���[�g���b�v
    Set mFile = Nothing
End Sub

Private Sub cmdOutMsg_Click()
    Dim NewEntryStartDate As String
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim cnt As Long
    
    NewEntryStartDate = Format(gdDBS.SystemUpdate("AANWDT"), "yyyy/mm/dd hh:nn:ss")
    sql = "SELECT COUNT(*) CNT "
    sql = sql & " FROM tfFurikaeYoteiData "
    sql = sql & " WHERE FASQNO = " & txtFurikaeBi.Number
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    cnt = dyn.Fields("CNT")
    Call dyn.Close
    Call pNormalEndMessage(True, cnt, NewEntryStartDate, vMsgMode:=True)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    Call mForm.LockedControl(False)
    lblMessage.Caption = mExeMsg
    'txtFurikaeBi.Number = gdDBS.SYSDATE("YYYYMMDD")
    txtFurikaeBi.Number = gdDBS.Nz(gdDBS.SystemUpdate("AANXKZ"))
    chkJisseki.Value = eCheckButton.Mukou  '�����ɐݒ�
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call gdDBS.Database.Rollback
    mAbort = True
    Set frmYoteiDataExport = Nothing
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

Private Sub txtFurikaeBi_DropOpen(NoDefault As Boolean)
    txtFurikaeBi.Calendar.Holidays = gdDBS.Holiday(txtFurikaeBi.Year)
End Sub

