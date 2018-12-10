VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBankDataImport 
   Caption         =   "���Z�@�փf�[�^�捞"
   ClientHeight    =   3345
   ClientLeft      =   2805
   ClientTop       =   1830
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5640
   Begin MSComctlLib.ProgressBar pgrProgressBar 
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "�捞(&I)"
      Height          =   435
      Left            =   2100
      TabIndex        =   2
      Top             =   2580
      Width           =   1395
   End
   Begin VB.CommandButton cmdRecv 
      Caption         =   "��M(&R)"
      Height          =   435
      Left            =   540
      TabIndex        =   1
      Top             =   2580
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "�I��(&X)"
      Height          =   435
      Left            =   3960
      TabIndex        =   0
      Top             =   2580
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   480
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   255
      Left            =   4020
      TabIndex        =   4
      Top             =   0
      Width           =   1395
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      Height          =   1635
      Left            =   480
      TabIndex        =   3
      Top             =   420
      Width           =   4815
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
Attribute VB_Name = "frmBankDataImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////////////////
'//�ǂ����Ă����p�E�S�p���݂̃g���~���O���o���Ȃ��̂ł�������.
Private Type tpBank         '//���Z�@��
    BankCode    As String * 4   '��s�R�[�h
    ShitenCode  As String * 3   '�x�X�R�[�h
    SeqCode     As String * 1   'SEQ-CODE       '��s=':#@,=' / �x�X='��`�','A�`Z','0�`9'
    KanaName    As String * 15  '��s��_�J�i
    KanjiName   As String * 30  '��s��_����
    HaitenInfo  As String * 4   '�p�X���       'Blank=�c�ƒ�,1�`9=�p�X��
    CrLf        As String * 2   'CR + LF
End Type

Private mCaption As String
Private Const mExeMsg As String = "��Ǝ菇" & vbCrLf & vbCrLf & "�@�P�F��M���������܂�." & vbCrLf & "�@�Q�F�捞���������܂�." & vbCrLf & vbCrLf & "�捞���ʂ��\������܂��̂œ��e�ɏ]���Ă�������." & vbCrLf & vbCrLf
Private mForm As New FormClass
Private mReg As New RegistryClass
Private mAbort As Boolean

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdImport_Click()
    Dim mFile As New FileClass
    
    dlgFile.DialogTitle = "�t�@�C�����J��(" & mCaption & ")"
    dlgFile.FileName = mReg.InputFileName(mCaption)
'//LZH �t�@�C���͉𓀂��Ă���̃X�^�[�g�Ƃ���B
#If 1 Then
    If IsEmpty(mFile.OpenDialog(dlgFile)) Then
        Exit Sub
    End If
#Else
'// �r���܂ŃR�[�f�B���O��������߂�.....�B
    If IsEmpty(mFile.OpenDialog(dlgFile, "LZḨ�� (*.lzh)|*.lzh")) Then
        Exit Sub
    End If
    
    '//�t�@�C�������h���C�u�`�g���q�܂ŕ���
    Dim drv As String, path As String, file As String, ext As String
'//2006/03/13 SplitPath() �Ƀo�O���������̂ŃR�����g���F�g�p���鎞�͂�������f�o�b�N���鎖�I
'    Call mFile.SplitPath(mReg.LzhExtractFile, drv, path, file, ext)
    '//�I�v�V����: e = Extract : ��
    '//�p�����[�^: -c ���t�`�F�b�N����
    '//            -m ���b�Z�[�W�}�~
    '//            -n �i���_�C�A���O��\��
    Dim ret As Integer, lzhMsg As String * 8192
    ret = Unlha(0, "e -c " & dlgFile.FileName & " " & (drv & path), lzhMsg, Len(lzhMsg))
#End If

    Dim mBank As tpBank, sql As String, SvrDate As String
    Dim updCnt As Long, insCnt As Long, delCnt As Long
    Dim fp As Integer
    Dim ms As New MouseClass
    Call ms.Start
    
    '//�X�V�O�̃T�[�o�[���t�擾
    SvrDate = gdDBS.sysDate
    mReg.InputFileName(mCaption) = dlgFile.FileName
    fp = FreeFile
    Open mReg.InputFileName(mCaption) For Random Access Read As fp Len = Len(mBank)
    pgrProgressBar.Max = LOF(fp) / Len(mBank)
    '//�t�@�C���T�C�Y���Ⴄ�ꍇ�̌x�����b�Z�[�W
    If pgrProgressBar.Max <> Int(pgrProgressBar.Max) Then
        If (LOF(fp) - 1) / Len(mBank) <> Int((LOF(fp) - 1) / Len(mBank)) Then
#If 1 Then
            '/�������s����Ƃc�a�����������Ȃ�̂Œ��~����
            Call gdDBS.MsgBox("�w�肳�ꂽ�t�@�C��(" & mReg.InputFileName(mCaption) & ")���ُ�ł��B" & vbCrLf & vbCrLf & "�����𑱍s�o���܂���B", vbCritical + vbOKOnly, mCaption)
            Exit Sub
#Else
            If vbOK <> gdDBS.MsgBox("�w�肳�ꂽ�t�@�C��(" & mReg.InputFileName(mCaption) & ")���ُ�ł��B" & vbCrLf & vbCrLf & "���̂܂ܑ��s���܂����H", vbInformation + vbOKCancel + vbDefaultButton2, mCaption) Then
                Exit Sub
            End If
#End If
        End If
    End If
    
    On Error GoTo cmdImport_ClickError
    Call gdDBS.Database.BeginTrans
    'Do Until EOF(fp)   '//���̍\�����ƍŏI���R�[�h�̎��܂œǍ��݂��Ă��܂��FEOF()�͂����������f
    Do While Loc(fp) < LOF(fp) / Len(mBank)
        DoEvents
        If mAbort Then
            GoTo cmdImport_ClickError
        End If
        Get fp, Loc(fp) + 1, mBank
        pgrProgressBar.Value = IIf(Loc(fp) <= pgrProgressBar.Max, Loc(fp), pgrProgressBar.Max)
        sql = "UPDATE tdBankMaster SET "
        sql = sql & " DAKJNM = '" & mFile.StrTrim(mBank.KanjiName) & "',"
        sql = sql & " DAKNNM = '" & mFile.StrTrim(mBank.KanaName) & "',"
        sql = sql & " DAHTIF = '" & mFile.StrTrim(mBank.HaitenInfo) & "',"
        sql = sql & " DAUPDT = SYSDATE"
        sql = sql & " WHERE DARKBN = '" & pGetRecordKubun(mBank.SeqCode) & "'"
        sql = sql & "   AND DABANK = '" & mFile.StrTrim(mBank.BankCode) & "'"
        sql = sql & "   AND DASITN = '" & mFile.StrTrim(mBank.ShitenCode) & "'"
        sql = sql & "   AND DASQNO = '" & mFile.StrTrim(mBank.SeqCode) & "'"
        If 0 <> gdDBS.Database.ExecuteSQL(sql) Then
            updCnt = updCnt + 1
        Else
            sql = "INSERT INTO tdBankMaster("
            sql = sql & "DARKBN,DABANK,DASITN,DASQNO,DAKNNM,DAKJNM,DAHTIF"
            sql = sql & ")VALUES("
            sql = sql & "'" & pGetRecordKubun(mBank.SeqCode) & "',"
            sql = sql & "'" & mFile.StrTrim(mBank.BankCode) & "',"
            sql = sql & "'" & mFile.StrTrim(mBank.ShitenCode) & "',"
            sql = sql & "'" & mFile.StrTrim(mBank.SeqCode) & "',"
            sql = sql & "'" & mFile.StrTrim(mBank.KanaName) & "',"
            sql = sql & "'" & mFile.StrTrim(mBank.KanjiName) & "',"
            sql = sql & "'" & mFile.StrTrim(mBank.HaitenInfo) & "'"
            sql = sql & ")"
            Call gdDBS.Database.ExecuteSQL(sql)
            insCnt = insCnt + 1
        End If
    Loop
    Close #fp
    '//�X�V�ΏۂłȂ��������R�[�h���폜����:�K���S������̂��O������I�I�I
    sql = "DELETE tdBankMaster "
    sql = sql & " WHERE DAUPDT < TO_DATE('" & Format(SvrDate, "yyyy-mm-dd hh:nn:ss") & "','yyyy-mm-dd hh24:mi:ss')"
    delCnt = gdDBS.Database.ExecuteSQL(sql)
    Dim AddMsg As String
    AddMsg = "�ǉ�=" & insCnt & ":�X�V=" & updCnt & ":�폜=" & delCnt & " ���̃f�[�^����荞�܂�܂����B"
    lblMessage.Caption = mExeMsg & AddMsg
    Call gdDBS.AutoLogOut(mCaption, AddMsg)
    
    Call gdDBS.Database.CommitTrans
    Exit Sub
cmdImport_ClickError:
    Call gdDBS.Database.Rollback
    Call gdDBS.ErrorCheck       '//�G���[�g���b�v
'// gdDBS.ErrorCheck() �̏�Ɉړ�
'//    Call gdDBS.Database.Rollback
    Call gdDBS.AutoLogOut(mCaption, " �G���[�������������ߎ捞�����͒��~����܂����B")
End Sub

Private Function pGetRecordKubun(ByVal vCode As Variant) As Integer
    pGetRecordKubun = Abs(vCode Like "[0-9]" Or vCode Like "[A-Z]" Or vCode Like "[�-�]")
End Function

Private Sub cmdRecv_Click()
    Call Shell(mReg.TransferCommand(mCaption), vbNormalFocus)
End Sub

Private Sub Form_Load()
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    lblMessage.Caption = mExeMsg
End Sub

Private Sub Form_Resize()
    Call mForm.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mAbort = True
    Set mForm = Nothing
    Set frmBankDataImport = Nothing
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

