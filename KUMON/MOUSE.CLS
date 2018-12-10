VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MouseClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Sub Start(Optional oPointer As Variant = vbArrowHourglass)
    Screen.MousePointer = oPointer
End Sub

Private Sub Class_Initialize()
    Screen.MousePointer = vbArrowHourglass
End Sub

Private Sub Class_Terminate()
    Screen.MousePointer = vbDefault
End Sub

