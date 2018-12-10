VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpreadClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mSpread As vaSpread
Attribute mSpread.VB_VarHelpID = -1
Private mOperationMode As Integer

Private Enum eSprRowFlag
    eDelete = -1
    eNonEdit = 0
    eEdit = 1
    eInsert = 2
    eEditHeader = 3
End Enum

Private mTopRow     As Long
Private mCurRow     As Long

Public Sub ComboBox(vCol As Long, vList As String, Optional vEditable As Boolean = False)
    With mSpread
        .Col = vCol
        .Row = -1
        .CellType = CellTypeComboBox
        .TypeComboBoxEditable = vEditable   '//�R���{�{�b�N�X���͕ҏW�s�H
        .TypeComboBoxList = vList
    End With
End Sub

Public Sub SavePosition()
    Me.Redraw = False
    mTopRow = mSpread.TopRow
    mCurRow = mSpread.ActiveRow
End Sub

Public Sub LoadPosition(Optional vColor As Long = vbCyan)
    Me.Redraw = True
    mSpread.TopRow = mTopRow
    Me.BackColor(-1, mCurRow) = vColor     '//�J�����g�s�ɐF�ݒ�
End Sub

Public Property Get RowDelete() As Variant:     RowDelete = eSprRowFlag.eDelete:   End Property
Public Property Get RowNonEdit() As Variant:    RowNonEdit = eSprRowFlag.eNonEdit: End Property
Public Property Get RowEdit() As Variant:       RowEdit = eSprRowFlag.eEdit:       End Property
Public Property Get RowInsert() As Variant:     RowInsert = eSprRowFlag.eInsert:   End Property
Public Property Get RowEditHeader() As Variant: RowEditHeader = eSprRowFlag.eEditHeader:   End Property

Public Sub Init(ByVal vNewSpread As vaSpread, Optional vBackColor As Long = vbButtonFace, Optional vLockBackColor As Long = 0&)
    Set mSpread = vNewSpread
    '//�������͓Ǎ��ݐ�p�Ƃ� Spread ���ǂ����N���b�N���������[�h��ύX����
    mSpread.OperationMode = OperationModeRead
    mOperationMode = OperationModeRead          '//�f�t�H���g�͓ǂݎ���p
    mSpread.GrayAreaBackColor = vBackColor
    If vLockBackColor Then
        mSpread.LockBackColor = &H80FFFF    'vbYellow
    End If
    mSpread.BackColorStyle = BackColorStyleUnderGrid    '//�O���b�h����\�����邽��
    mSpread.ShadowDark = mSpread.ShadowColor            '//�w�b�_�[�̉e��\�ʂƓ����ɂ��ĂR�c��a�炰��
    '//���׍s���󔒉�
    Dim Rows As Long
    Rows = mSpread.MaxRows
    mSpread.MaxRows = 0
    mSpread.MaxRows = Rows
End Sub

Public Sub OddEvenRowColor(Optional vBackOdd As Variant, Optional vForeOdd As Variant, Optional vBackEven As Variant, Optional vForeEven As Variant)
    Dim clrBackOdd  As Long     '��s�̔w�i�F
    Dim clrForeOdd  As Long     '��s�̃e�L�X�g�F
    Dim clrBackEven As Long     '�����s�̔w�i�F
    Dim clrForeEven As Long     '�����s�̃e�L�X�g�F
    Call mSpread.GetOddEvenRowColor(clrBackOdd, clrForeOdd, clrBackEven, clrForeEven)
    If Not IsMissing(vBackOdd) Then '��s�̔w�i�F
        clrBackOdd = vBackOdd
    End If
    If Not IsMissing(vForeOdd) Then    '��s�̃e�L�X�g�F
        clrForeOdd = vForeOdd
    End If
    If Not IsMissing(vBackEven) Then    '�����s�̔w�i�F
        clrBackEven = vBackEven
    End If
    If Not IsMissing(vForeEven) Then    '�����s�̃e�L�X�g�F
        clrForeEven = vForeEven
    End If
    Call mSpread.SetOddEvenRowColor(clrBackOdd, clrForeOdd, clrBackEven, clrForeEven)
End Sub

#If 0 Then
Public Sub InitialCopy(vSource As vaSpread)
#If 1 Then
'//�Z���u���b�N�R�s�[�ł͑������R�s�[�ł��Ȃ��H
    Dim file As New FileClass
    Dim fname As String
    fname = file.MakeTempFile
    Call vSource.SaveToFile(fname, False)
    Call mSpread.LoadFromFile(fname)
    Set file = Nothing  '//�ꎞ�t�@�C�����폜������
#Else
    With vSource
        .Col = 1
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = .MaxRows
        .Action = ActionSelectBlock
        .Action = ActionClipboardCopy
    End With
    With mSpread
        .Col = 1
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = .MaxRows
        .Action = ActionSelectBlock
        .Action = ActionClipboardPaste
    End With
#End If
End Sub
#End If

Public Property Get BackColor(ByVal vCol As Long, ByVal vRow As Long) As Long
    mSpread.Col = vCol
    mSpread.Row = vRow
    BackColor = mSpread.BackColor
End Property

Public Property Let BackColor(ByVal vCol As Long, ByVal vRow As Long, vColor As Long)
#If 0 Then  '//2006/03/09 ���Z�b�g���R�����g���F�s�s��������Ό��ɖ߂�����
    mSpread.Col = -1
    mSpread.Row = -1
    mSpread.BackColor = vbWhite
#End If
    mSpread.Col = vCol
    mSpread.Row = vRow
    mSpread.BackColor = vColor
End Property

Public Property Get ForeColor(ByVal vCol As Long, ByVal vRow As Long) As Long
    mSpread.Col = vCol
    mSpread.Row = vRow
    ForeColor = mSpread.ForeColor
End Property

Public Property Let ForeColor(ByVal vCol As Long, ByVal vRow As Long, vColor As Long)
    mSpread.Col = vCol
    mSpread.Row = vRow
    mSpread.ForeColor = vColor
End Property

Public Sub Sort(ByVal vKey1 As Long, _
       Optional ByVal vOrder1 As Long = fpSpread.SortKeyOrderAscending, _
       Optional ByVal vKey2 As Long = 0, _
       Optional ByVal vOrder2 As Long = fpSpread.SortKeyOrderAscending, _
       Optional ByVal vKey3 As Long = 0, _
       Optional ByVal vOrder3 As Long = fpSpread.SortKeyOrderAscending _
    )
    '//��肠�����u�s�̕��ёւ��v���R���...�B
    With mSpread
        .Row = 1
        .Col = 1
        .Row2 = .MaxRows
        .Col2 = .MaxCols
        .SortBy = SortByRow
        .SortKey(1) = vKey1
        .SortKeyOrder(1) = vOrder1
        If vKey2 Then
            .SortKey(2) = vKey2
            .SortKeyOrder(2) = vOrder2
        End If
        If vKey3 Then
            .SortKey(3) = vKey3
            .SortKeyOrder(3) = vOrder3
        End If
        .Action = ActionSort
    End With
End Sub

Public Property Get Locked(ByVal vCol As Long, ByVal vRow As Long) As Boolean
    With mSpread
        .Col = vCol
        .Row = vRow
        Locked = .Lock
    End With
End Property

Public Property Let Locked(ByVal vCol As Long, ByVal vRow As Long, ByVal vNewValue As Boolean)
    With mSpread
        .Col = vCol
        .Row = vRow
        .Lock = vNewValue
    End With
End Property

Public Property Let BlockLocked(ByVal vCol As Long, ByVal vRow As Long, ByVal vCol2 As Long, ByVal vRow2 As Long, ByVal vLock As Boolean)
    With mSpread
        .BlockMode = True
        .Col = vCol
        .Col2 = vCol2
        .Row = vRow
        .Row2 = vRow2
        .Lock = vLock
        .BlockMode = False
    End With
End Property

Public Sub AddRow(Optional ByVal vRow As Long = 0)
    With mSpread
        If vRow = 0 Or vRow > .MaxRows Then
            .MaxRows = .MaxRows + 1
        Else
            .Row = vRow
            .Row2 = vRow
            .BlockMode = True
            .Action = ActionInsertRow
            .BlockMode = False
        End If
    End With
End Sub

Public Function AddMaxRow(vKeyCode As Integer, vCheck As Boolean, vMaxCol As Long) As Boolean
    '//�ŏI�s�� vbKeyDown / vbKeyReturn �����������Ƃ�
    If vKeyCode = vbKeyDown Then
        '//�ŏI�s�� vbKeyDown �����������Ƃ�
        If mSpread.ActiveRow < mSpread.MaxRows Then
            Exit Function
        End If
    ElseIf vKeyCode = vbKeyReturn Then
        '//�ŏI�s�� vbKeyDown �����������Ƃ�
        If mSpread.ActiveRow < mSpread.MaxRows _
        Or mSpread.ActiveCol < vMaxCol Then
            Exit Function
        End If
    Else
        Exit Function
    End If
    Dim Row As Long
    Row = mSpread.MaxRows
    '//�R�[�h�����͂���Ă���΍s��ǉ�
'    If "" <> Trim(Me.Value(vCheckCol, mSpread.MaxRows)) Then
    If vCheck Then
        Call Me.AddRow
    End If
    AddMaxRow = Row < mSpread.MaxRows '//�s�𑝂₹���̂� True ��Ԃ�.
End Function

Public Sub ActiveCell(ByVal vCol As Long, ByVal vRow As Long)
    mSpread.Col = vCol
    mSpread.Row = vRow
    mSpread.Action = ActionActiveCell
End Sub

Public Property Get Text(ByVal vCol As Long, ByVal vRow As Long) As Variant
    '//�R���{�{�b�N�X�̑I�𕶎��͂��̕��@�Ŏ擾�FListIndex �� Value ��
    With mSpread
        .Col = vCol
        .Row = vRow
        Text = .Text
    End With
End Property

Public Property Let Text(ByVal vCol As Long, ByVal vRow As Long, ByVal vNewText As Variant)
    '//�R���{�{�b�N�X�̑I�𕶎��͂��̕��@�Ŏ擾�FListIndex �� Value ��
    With mSpread
        .Col = vCol
        .Row = vRow
        .Text = vNewText
    End With
End Property

Public Property Get Value(ByVal vCol As Long, ByVal vRow As Long) As Variant
    With mSpread
        .Col = vCol
        .Row = vRow
        Value = .Value
    End With
End Property

Public Property Let Value(ByVal vCol As Long, ByVal vRow As Long, ByVal vNewValue As Variant)
    With mSpread
        .Col = vCol
        .Row = vRow
        .Value = vNewValue
    End With
End Property

Public Sub LineDelete(vFlagCol As Long)
    If Me.BackColor(-1, mSpread.ActiveRow) = vbRed Then
        Me.BackColor(-1, mSpread.ActiveRow) = vbWhite
        Me.Value(vFlagCol, mSpread.ActiveRow) = Me.RowNonEdit
    Else
        Me.BackColor(-1, mSpread.ActiveRow) = vbRed
        Me.Value(vFlagCol, mSpread.ActiveRow) = Me.RowDelete
    End If
End Sub

#If 1 Then
Public Property Let ClipValue(vCol As Long, vRow As Long, vCol2 As Long, vRow2 As Long, vTabData As String)
    With mSpread
        .Col = vCol
        .Row = vRow
        .Col2 = vCol2
        .Row2 = vRow2
        .ClipValue = vTabData
    End With
End Property
#Else
Public Sub ClipValue(vTabStr As Variant, _
            Optional vCol As Long = 1, _
            Optional vRow As Long = 1, _
            Optional vCol2 As Long = 0, _
            Optional vRow2 As Long = 0, _
            Optional vMaxRows As Long = 0 _
        )
    With mSpread
        .MaxRows = 0
        .MaxRows = vMaxRows
        .Col = vCol
        .Row = vRow
        .Col2 = IIf(vCol2, vCol2, .MaxCols)
        .Row2 = IIf(vRow2, vRow2, .MaxRows)
        .ClipValue = vTabStr
    End With
End Sub
#End If

Public Property Get MaxCols() As Long
    MaxCols = mSpread.MaxCols
End Property
    
Public Property Let MaxCols(vCols As Long)
    mSpread.MaxCols = vCols
End Property
    
Public Property Get MaxRows() As Long
    MaxRows = mSpread.MaxRows
End Property
    
Public Property Let MaxRows(vRows As Long)
    mSpread.MaxRows = vRows
End Property
    
Public Property Let OperationMode(vMode As OperationModeConstants)
    Dim blocks As Integer
    blocks = mSpread.SelectBlockOptions  '//�Z���u���b�N�I��ێ�
    mOperationMode = vMode
    mSpread.SelectBlockOptions = blocks '//�Z���u���b�N�I�𕜋A
End Property

Private Sub mSpread_Click(ByVal Col As Long, ByVal Row As Long)
    Dim blocks As Integer
    blocks = mSpread.SelectBlockOptions  '//�Z���u���b�N�I��ێ�
    mSpread.OperationMode = mOperationMode
    mSpread.SelectBlockOptions = blocks '//�Z���u���b�N�I�𕜋A
End Sub

Public Property Let ColWidth(ByVal vCol As Long, vWidth As Currency)
    mSpread.ColWidth(vCol) = vWidth
End Property

Public Property Let RowHeight(ByVal vRow As Long, vHeight As Currency)
    mSpread.RowHeight(vRow) = vHeight
End Property

Public Sub Refresh()
    Call mSpread.Refresh
End Sub

Public Property Let Redraw(vMode As Boolean)
    mSpread.Redraw = vMode
End Property

Public Sub CellType(ByVal vCol As Long, vType As Integer, _
            Optional vFormat As Variant = "", _
            Optional vMin As Variant = "", _
            Optional vMax As Variant = "")
    With mSpread
        .Col = vCol
        .Row = -1
        .CellType = vType
        Select Case .CellType
        Case CellTypeStaticText '//Label
            .TypeTextWordWrap = UCase("WordWrap") = UCase(vFormat)     '//vFormat="WordWrap" => �s�܂�Ԃ�
            .TypeHAlign = Val(vMax)
        Case CellTypeEdit       '//Text
            If vFormat <> "" Then
                .TypeMaxEditLen = Len(vFormat)
            End If
            .TypeHAlign = Val(vMin)
        Case CellTypeInteger    '//����
            .TypeIntegerMin = Val(vMin)
            .TypeIntegerMax = Val(vMax)
        Case CellTypeFloat      '//����
            .TypeFloatMin = Val(vMin)
            .TypeFloatMax = Val(vMax)
            If InStr(vMax, ".") Then
                .TypeFloatDecimalPlaces = Len(vMax) - InStr(vMax, ".")
            Else
                .TypeFloatDecimalPlaces = 0
            End If
        Case CellTypeDate       '//���t
        Case CellTypeTime       '//����
        Case Else               '//���͏���ɒ�`�����I
        End Select
    End With
End Sub

Public Sub FindSetColor(vCol As Long, vFind As String, vColor As Long)
    Dim ixStart As Long, ix As Long
    With mSpread
        ixStart = .RowsFrozen + 1
        For ix = ixStart To .MaxRows
            If 0 < InStr(Me.Value(vCol, ix), vFind) Then
                .Col = -1
                .Row = ix
                .BackColor = vColor
            End If
        Next ix
    End With
End Sub

Public Sub GroupMask(vCol As Long)
    Dim ixStart As Long, ix0 As Long, ix1 As Long
    With mSpread
        ixStart = .RowsFrozen + 1
        For ix0 = ixStart To .MaxRows
#If 1 Then
            ix1 = ix0 + 1
            Do While Value(vCol, ix0) = Value(vCol, ix1)
                Value(vCol, ix1) = ""
                ix1 = ix1 + 1
                If ix1 > .MaxRows Then
                    Exit Do
                End If
            Loop
#Else
            For ix1 = ix0 + 1 To .MaxRows
                If Value(vCol, ix0) = Value(vCol, ix1) Then
                    Value(vCol, ix1) = ""
                Else
                    Exit For
                End If
            Next ix1
#End If
            ix0 = ix1 - 1
        Next ix0
    End With
End Sub

Public Function ExportRangeToText(vCol As Long, vRow As Long, vCol2 As Long, vRow2 As Long, vFileName As String, vDelimiter As String) As Boolean
    '// Spread �֐��ł͂ǂ����w�b�_�[���o�͂ł��Ȃ��̂ō쐬
    Dim fp As Integer
    Dim Col As Long, Row As Long, tmpStr As String
    Dim ms As New MouseClass
    
    Call ms.Start
    fp = FreeFile
    Open vFileName For Output As #fp
    For Row = vRow To vRow2
        tmpStr = ""
        For Col = vCol To vCol2
            '==> �"� �t�ł��Ă�Excel�œǍ��ݎ��͓����Ȃ̂Ńo�C�g�������Ȃ��ق���...�B
            'tmpStr = tmpStr & """" & Me.Value(Col, Row) & ""","
            tmpStr = tmpStr & Me.Value(Col, Row) & vDelimiter
        Next Col
        Print #fp, Left(tmpStr, Len(tmpStr) - Len(vDelimiter))
    Next Row
    Close #fp
    ExportRangeToText = True
    Exit Function
ExportRangeToTextError:
End Function
