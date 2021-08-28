VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Binary
Option Explicit

Private Sub cbNewGame_Click()
    Call ClearField
    Call PrepareField
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Intersect(Target, Range(Cells(1, 1), Cells(16, 32))) Is Nothing Then
        Module1.bSelectionChanged = True
        Module1.ScheduleAction
    End If
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, Range(Cells(1, 1), Cells(16, 32))) Is Nothing Then
        Cancel = True
        Module1.bRightClicked = True
        If Not Module1.bSelectionChanged Then
            Module1.ScheduleAction
        End If
    End If
End Sub




