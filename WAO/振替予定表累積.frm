VERSION 5.00
Begin VB.Form frmFurikaeDataRuiseki 
   Caption         =   "�U�֗\��\ �� ���ʒm��(�ݐ�)"
   ClientHeight    =   3255
   ClientLeft      =   2295
   ClientTop       =   2130
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6045
   Begin VB.ListBox lstFurikaeBi 
      Height          =   690
      ItemData        =   "�U�֗\��\�ݐ�.frx":0000
      Left            =   2340
      List            =   "�U�֗\��\�ݐ�.frx":0016
      Style           =   1  '�����ޯ��
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "���s(&E)"
      Height          =   435
      Left            =   540
      TabIndex        =   1
      Top             =   2580
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "�I��(&X)"
      Default         =   -1  'True
      Height          =   435
      Left            =   4140
      TabIndex        =   0
      Top             =   2580
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "�����U�֓�"
      Height          =   195
      Left            =   1260
      TabIndex        =   4
      Top             =   1740
      Width           =   915
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   0
      Width           =   1395
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      Height          =   1155
      Left            =   360
      TabIndex        =   2
      Top             =   420
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
Attribute VB_Name = "frmFurikaeDataRuiseki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCaption As String
Private mForm As New FormClass

Private Const mExeMsg As String = "�U�֋��z�\��\ �� ���ʒm���f�[�^��ݐς��܂��B" & vbCrLf & vbCrLf

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdExec_Click()
    Dim sql As String, cnt As Long
    Dim WhereSQL As String, ix As Long, msg As String
    Dim ms As New MouseClass
    Call ms.Start
    
'//2004/05/17 �ݐό����̃��O�ǉ��̂��߂ɓ��t��ޔ�
    Dim RuisekiDate As String
    '//���X�g�Ń`�F�b�N���ꂽ�f�[�^�� IN ���...�B
    WhereSQL = " WHERE FASQNO IN("
    RuisekiDate = "("
    For ix = 0 To lstFurikaeBi.ListCount - 1
        If lstFurikaeBi.Selected(ix) = True Then
            cnt = cnt + 1
            WhereSQL = WhereSQL & Format(lstFurikaeBi.List(ix), "yyyymmdd") & ","
'//2004/05/17 �ݐό����̃��O�ǉ��̂��߂ɓ��t��ޔ�
            RuisekiDate = RuisekiDate & lstFurikaeBi.List(ix) & ","
        End If
    Next ix
    WhereSQL = Left(WhereSQL, Len(WhereSQL) - 1) & ")"
'//2004/05/17 �ݐό����̃��O�ǉ��̂��߂ɓ��t��ޔ�
    RuisekiDate = Left(RuisekiDate, Len(RuisekiDate) - 1) & ")"
    If cnt = 0 Then
        msg = "�ݐς��ׂ��f�[�^�͂���܂���ł����B"
        lblMessage.Caption = mExeMsg & msg
        Call MsgBox(msg, vbInformation, mCaption)
        Exit Sub
    End If
    
    On Error GoTo cmdExec_ClickError

'//2003/02/03 �X�V��ԃt���O���`�F�b�N���Čx�� �ǉ�:0=DB�쐬,1=�\��쐬,2=�\��捞,3=�����쐬
    Dim dyn As Object
    sql = "SELECT FASQNO FROM tfFurikaeYoteiData"
    sql = sql & WhereSQL
    sql = sql & " AND FAUPFG < '" & eKouFuriKubun.SeikyuText & "'"
    sql = sql & " AND NVL(FAKYFG,0) = 0"
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
    If dyn.RecordCount > 0 Then
        msg = "�����f�[�^�̍쐬����Ă��Ȃ��f�[�^�����݂��܂�." & vbCrLf & "�ݐϏ����𑱍s���܂����H"
        lblMessage.Caption = mExeMsg & msg
        If vbOK <> MsgBox(msg, vbInformation + vbOKCancel + vbDefaultButton2, mCaption) Then
            Exit Sub
        End If
    End If
    
    Call gdDBS.Database.BeginTrans
    
'//2004/06/03 �V�K�����������t��ی�҃}�X�^(CANWDT)�ɐݒ�F���z�u�O�v�͐V�K�Ƃ��Ȃ�
    sql = "UPDATE tcHogoshaMaster SET "
    sql = sql & " CANWDT = SYSDATE "
    sql = sql & " WHERE (CAITKB,CAKYCD,CAHGCD) IN("
    sql = sql & "       SELECT FAITKB,FAKYCD,FAHGCD "
    sql = sql & "       FROM tfFurikaeYoteiData "
    '//2007/04/19 WAO�͋��z���͖����Ȃ̂ŏ������͂���
    'sql = sql & "       WHERE (NVL(faskgk,0) > 0 OR NVL(fahkgk,0) > 0) "
    sql = sql & "     )"
    sql = sql & "  AND CANWDT IS NULL"
    cnt = gdDBS.Database.ExecuteSQL(sql)
    
    '//�ݐ�
    sql = "INSERT INTO tfFurikaeYoteiTran "
    sql = sql & " SELECT * FROM tfFurikaeYoteiData"
    sql = sql & WhereSQL
    cnt = gdDBS.Database.ExecuteSQL(sql)
    '//�ݐς��������폜
    sql = " DELETE tfFurikaeYoteiData"
    sql = sql & WhereSQL
    Call gdDBS.Database.ExecuteSQL(sql)
'//2003/02/04 ����U�����E��������U�֓� ���X�V����
    Dim KouFuriDay As Integer, FurikomiDay As Integer
    Dim KouFuriDate As Date, FurikomiDate As Date

    '//�U�����F�_��҈���
    FurikomiDay = gdDBS.SystemUpdate("AAFKDT")
    '//�����̐U�������Z�o�����
    FurikomiDate = DateSerial( _
                        Mid(gdDBS.SystemUpdate("AANXFK"), 1, 4), _
                        Mid(gdDBS.SystemUpdate("AANXFK"), 5, 2) + 1, _
                        FurikomiDay _
                    )
    '//����U���� �ݒ�
    gdDBS.SystemUpdate("AANXFK") = Format(NextDay(FurikomiDate), "yyyymmdd")
    
    '//�����U�֓��F�ی�҈���
    KouFuriDay = gdDBS.SystemUpdate("AAKZDT")
    '//�����̌����U�֓����Z�o�����
'//2010/02/23 �Q�O�P�O�N�Q���� 2/27,28 ���c�Ɠ��łȂ��ׁA�U�֓��� 3/1 �ɂȂ��Ă��܂��Ă���̂łP�������ݒ肵�Ă��܂��o�O�Ή�
    Dim wDay As Integer, addMonth As Integer
    wDay = Right(gdDBS.SystemUpdate("AANXKZ"), 2)
    If KouFuriDay <= wDay Then
        addMonth = 1
    End If
    KouFuriDate = DateSerial( _
                        Mid(gdDBS.SystemUpdate("AANXKZ"), 1, 4), _
                        Mid(gdDBS.SystemUpdate("AANXKZ"), 5, 2) + addMonth, _
                        KouFuriDay _
                    )
    KouFuriDate = NextDay(KouFuriDate)
    '//��������U�֓� �ݒ�
    gdDBS.SystemUpdate("AANXKZ") = Format(NextDay(KouFuriDate), "yyyymmdd")

'//2004/04/12 �����U�֓����r���Ĉȍ~�̓��� �Đݒ�
    If FurikomiDate < KouFuriDate Then      '//�N����
        If FurikomiDay < KouFuriDay Then    '//�@�@��
            FurikomiDate = DateSerial( _
                                Mid(gdDBS.SystemUpdate("AANXKZ"), 1, 4), _
                                Mid(gdDBS.SystemUpdate("AANXKZ"), 5, 2) + 1, _
                                FurikomiDay _
                            )
        Else
            FurikomiDate = DateSerial( _
                                Mid(gdDBS.SystemUpdate("AANXKZ"), 1, 4), _
                                Mid(gdDBS.SystemUpdate("AANXKZ"), 5, 2) + 0, _
                                FurikomiDay _
                            )
        End If
        '//����U���� �Đݒ�
        gdDBS.SystemUpdate("AANXFK") = Format(NextDay(FurikomiDate), "yyyymmdd")
    End If
'//2004/05/17 �ݐό����̃��O�ǉ�
    Call gdDBS.AutoLogOut(mCaption, "�����U�ւc�a�ݐ� = " & cnt & " �� �Ώ� = " & RuisekiDate)

    lblMessage.Caption = mExeMsg & cnt & " ���̃f�[�^���ݐς���܂����B"
    '//���s�X�V�t���O�ݒ�
    gdDBS.SystemUpdate("AAUPDE") = 1
    Call gdDBS.Database.CommitTrans
    Exit Sub
cmdExec_ClickError:
    Call gdDBS.Database.Rollback
    Call gdDBS.ErrorCheck       '//�G���[�g���b�v
'// gdDBS.ErrorCheck() �̏�Ɉړ�
'//    Call gdDBS.Database.Rollback
End Sub

Private Function NextDay(vStartDate As Variant) As Variant
    Dim ix As Integer
    Dim dyn As Object, sql As String
    '//�Q�O�A�x�͖������낤!!!
    For ix = 0 To 20
        NextDay = DateSerial(Year(vStartDate), Month(vStartDate), Day(vStartDate) + ix)
        '//1=���j��,2=���j��...,7=�y�j�� �Ȃ̂łQ�ȏ�͌��j��������j���̂͂�
        If (Weekday(NextDay, vbSunday) Mod 7) >= 2 Then
            sql = "SELECT EADATE FROM teHolidayMaster "
            sql = sql & " WHERE EADATE = " & Format(NextDay, "yyyymmdd")
            Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
            If dyn.EOF() Then
                Exit Function
            End If
            Call dyn.Close
        End If
    Next ix
    '//�I�[�o�[�����̂�...�B
    NextDay = vStartDate
End Function

Private Sub Form_Load()
    Dim reg As New RegistryClass
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    lblMessage.Caption = mExeMsg
    
    '//ListBox �Ɍ��݂̗\���S�ă��X�g�A�b�v����B
'    Dim sql As String, dyn As OraDynaset
    Dim sql As String, dyn As Object
    sql = "SELECT FASQNO,TO_CHAR(TO_DATE(FASQNO,'YYYYMMDD'),'YYYY/MM/DD') AS FaDate"
    sql = sql & " FROM tfFurikaeYoteiData"
    sql = sql & " GROUP BY FASQNO"
    sql = sql & " ORDER BY FASQNO"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    Call lstFurikaeBi.Clear
    Do Until dyn.EOF()
        Call lstFurikaeBi.AddItem(dyn.Fields("FaDate"))
'        lstFurikaeBi.Selected(lstFurikaeBi.NewIndex) = True
        Call dyn.MoveNext
    Loop
    Call dyn.Close
#If 0 Then
'''//�`�F�b�N�{�b�N�X�e�X�g�̂��߂ɍ쐬
    Dim i As Integer
    For i = 1 To 10
        Call lstFurikaeBi.AddItem(Format(Now() + i, "yyyy/mm/dd"))
        lstFurikaeBi.Selected(lstFurikaeBi.NewIndex) = True
    Next i
#End If
    cmdExec.Enabled = lstFurikaeBi.ListCount > 0
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mForm = Nothing
    Set frmFurikaeDataRuiseki = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub lstFurikaeBi_ItemCheck(Item As Integer)
    '//�`�F�b�N�{�b�N�X�͏�Ƀ`�F�b�N��ԂɈێ�����
'    lstFurikaeBi.Selected(Item) = True
End Sub

Private Sub mnuEnd_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuVersion_Click()
    Call frmAbout.Show(vbModal)
End Sub

