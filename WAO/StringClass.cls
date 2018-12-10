VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'   Function LenMbcs (ByVal str as String)
'      LenMbcs = LenB(StrConv(str, vbFromUnicode))
'   End Function
'
'   Dim MyString, MyLen
'   MyString = "ＡＢc"
'   "Ａ" と "Ｂ" は全角文字で "c" は半角文字です。
'   MyLen = Len(MyString)
'   ' 文字数として 3 が返されます。
'   MyLen = LenB(MyString)      Windows の場合は 6 がバイト数として返されます。
'   MyLen = LenMbcs(MyString)   Windows の場合は 5 が返されます。

Public Property Get CheckLength(vStr As String, Optional vCode As Variant = vbWide) As Boolean
    If vCode = vbWide Then
        CheckLength = LenB(StrConv(vStr, vbFromUnicode)) = LenB(StrConv(vStr, vbWide))
    Else
        CheckLength = Hankaku(vStr) 'LenB(vStr) = LenB(StrConv(vStr, vCode))
    End If
End Property

'Private Property Get Zenkaku(vStr As String) As Boolean
'    Dim idx As Integer
'    For idx = 1 To Len(vStr)
'
'    Next idx
'End Property

Private Property Get Hankaku(vStr As String) As Boolean
    Dim idx As Integer
    For idx = 1 To Len(vStr)
        If Asc(Mid(vStr, idx, 1)) < 0 And &HFF < Asc(Mid(vStr, idx, 1)) Then
            Exit Property
        End If
    Next idx
    Hankaku = True
End Property

Public Property Get FixedFormat(vNum As Variant, vLength As Integer) As String
    On Error Resume Next
    FixedFormat = Right(Space(vLength) & Format(vNum, "#,0"), vLength)
End Property
