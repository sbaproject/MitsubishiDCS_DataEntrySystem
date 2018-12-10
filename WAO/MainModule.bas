Attribute VB_Name = "MainModule"
Option Explicit
Public gdDBS  As DatabaseClass
Public gdForm As Form
Public gdFormSub As Form         '//�q���̉�ʂ����݂���H

'//2006/03/10 �ی�҃}�X�^�̎捞���[�U�[�h�c
Public Const gcImportUserName As String = "PUNCH_IMPORT"
Public Const gcTsuchoKigoMinLen As Integer = 3
Public Const gcTsuchoBangoMinLen As Integer = 7
Public Const gcFurikaeImportToDelete As String = "D"    '//�\��p��
Public Const gcImportHogoshaUser    As String = "$KZ_IMP"   '//�����U�ֈ˗����E�捞���[�U�[
Public Const gcFurikaeImportToMaster As String = "M"    '//�}�X�^�[���f

Public Enum eBankKubun
    KinnyuuKikan = 0
    YuubinKyoku = 1
End Enum

Public Enum eBankRecordKubun
    Bank = 0
    Shiten = 1
End Enum

Public Enum eBankYokinShubetsu
    Dummy = 0
    Futsuu = 1
    Touza = 2
End Enum

Public Enum eShoriKubun
    Add = 0
    Edit = 1
    Delete = 2
    Refer = 3   '//2012/12/07 �Q�Ƃ̃I�v�V�����{�^���ǉ�
End Enum

Public Enum eKouFuriKubun
    YoteiDB = 0         '//�\��c�a�쐬
    YoteiText = 1       '//�\��e�L�X�g�쐬
    YoteiImport = 2     '//�\��f�[�^�捞
    SeikyuText = 3      '//�����e�L�X�g�쐬
End Enum

#If 0 Then
'//�S��}�X�^��荞�ݗp
Declare Function Unlha Lib "Unlha32.DLL" (ByVal hWnd As Integer, ByVal szCmdline As String, ByVal szOutPutMsg As String, ByVal dwSize As Long) As Integer
#End If

'//2014/06/11 ���X�g����I�ׂ���e���`�F�b�N���邽�߂ɒ萔��
Public Const cKAIYAKU_DATA As String = "�ی�҃}�X�^�͉���Ԃł�."
Public Const cEXISTS_DATA As String = "�ی�҃}�X�^�Ɋ��ɑ��݂��܂�."

Sub Main()
    Dim mFile As New FileClass, path As String, drv As String
    Set gdDBS = New DatabaseClass
    Call frmMainMenu.Show
End Sub

Sub gkAllEnd()
    Set gdDBS = Nothing
    Set gdForm = Nothing
    End
End Sub


'////////////////////////////////////////////////////////////////////
'//EXE �̃v���O�����z���� \Backup �t�H���_���쐬���ăo�b�N�A�b�v����
Public Function gBackupTextData(vFileName As String) As Boolean
    Dim mFile As New FileClass
    Dim dstPath As String, dstDrv As String
    Dim dstFile As String, dstExt As String
    Call mFile.SplitPath(App.path, vDrv:=dstDrv, vPath:=dstPath, vMode:=True)
    Call mFile.SplitPath(vFileName, vFile:=dstFile, vExt:=dstExt)
    On Error Resume Next
    If "" = Dir(dstDrv & dstPath & "\Backup\") Then
        Call MkDir(dstDrv & dstPath & "\Backup")
    End If
    Call FileCopy(vFileName, dstDrv & dstPath & "\Backup\" & Format(Now(), "yyyymmdd.hhnnss") & "." & dstFile & dstExt)
End Function

