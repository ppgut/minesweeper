Attribute VB_Name = "Module1"
Option Explicit

Public bRightClicked As Boolean
Public bSelectionChanged As Boolean

Const MineSign As String = "*"
Dim NumberOfMines As Long

Dim lUnhidenFieldsCount As Long
Dim lMarkedFiledWOMine As Long
Dim arrField(1 To 16, 1 To 32) As Variant
Dim bLost As Boolean
Dim bWon As Boolean

Sub PrepareField()

    Call GenerateMines
    Call GenerateCounters
    
    Sheet1.Range(Cells(1, 1), Cells(16, 32)).Value = arrField
End Sub

Sub ClearField()
    Dim rField As Range
    Set rField = Range(Cells(1, 1), Cells(16, 32))
    
    With rField.Cells
        .ClearContents
        .Interior.Color = rgbLightGrey
        .Borders.LineStyle = xlContinuous
        .BorderAround xlContinuous, xlThick
        .NumberFormat = ";;;"
    End With
    lUnhidenFieldsCount = 0
    lMarkedFiledWOMine = 0
    bLost = False
    bWon = False
    Application.EnableEvents = True
End Sub

Sub GenerateMines()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Erase arrField
    
    If Not IsNumeric(Range("AL6").Value) Then
        Range("AL6").Value = 99
    ElseIf Range("AL6").Value < 25 Then
        Range("AL6").Value = 25
    ElseIf Range("AL6").Value > 99 Then
        Range("AL6").Value = 99
    End If
    
    NumberOfMines = CLng(Range("AL6").Value)
        
    For k = 1 To NumberOfMines
        Randomize
        i = Int((UBound(arrField, 1) * Rnd) + 1)
        j = Int((UBound(arrField, 2) * Rnd) + 1)
        If arrField(i, j) <> MineSign Then
            arrField(i, j) = MineSign
        Else
            k = k - 1
        End If
    Next k

End Sub

Sub GenerateCounters()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim m As Integer
    Dim n As Integer
    
    For i = 1 To UBound(arrField, 1)
        For j = 1 To UBound(arrField, 2)
            If arrField(i, j) <> MineSign Then
                
                n = 0
                For k = fMax(i - 1, 1) To fMin(i + 1, UBound(arrField, 1))
                    For m = fMax(j - 1, 1) To fMin(j + 1, UBound(arrField, 2))
                        If arrField(k, m) = MineSign Then
                            n = n + 1
                        End If
                    Next m
                Next k
                If n > 0 Then arrField(i, j) = n
                
            End If
        Next j
    Next i
    
End Sub

Private Function fMax(v1 As Integer, v2 As Integer) As Integer
    fMax = v1
    If v2 > v1 Then fMax = v2
End Function
Private Function fMin(v1 As Integer, v2 As Integer) As Integer
    fMin = v1
    If v2 < v1 Then fMin = v2
End Function

Public Sub UnHide(ByVal rng As Range, Optional bClicked As Boolean = False)
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim m As Integer
    
    i = rng.Row
    j = rng.Column
    
    If rng.Interior.Color <> vbWhite And rng.Interior.Color <> vbRed Then
        rng.Interior.Color = vbWhite
        rng.NumberFormat = "@"
        lUnhidenFieldsCount = lUnhidenFieldsCount + 1
        Select Case rng.Value
        Case ""
            For k = -1 To 1
                For m = -1 To 1
                    If i + k >= 1 And i + k <= 16 And j + m >= 1 And j + m <= 32 Then
                        If k <> i Or m <> j Then
                            'crashes if too many fields are to be unhidden (too deep recursion) - can be tested with amount of mines set to 1
                            UnHide rng.Offset(k, m)
                        End If
                    End If
                Next m
            Next k
        Case MineSign
            If bClicked Then
                bLost = True
                MsgBox "(x_x)"
                Call ClearField
                Call PrepareField
            End If
        Case Else
        End Select
    End If

End Sub

Sub Win()
    If Not bLost And Not bWon And lUnhidenFieldsCount + NumberOfMines = 16 * 32 Then
        bWon = True
        MsgBox "(^.^)"
    End If
End Sub

Sub ScheduleAction()
    Application.OnTime Now + TimeSerial(0, 0, 0.1), "DoAction"
End Sub

Sub DoAction()
'    Debug.Print "bRightClicked=" & bRightClicked
'    Debug.Print "bSelectionChanged=" & bSelectionChanged
    
    Selection.Cells(1, 1).Select
    
    If bRightClicked Then
        Select Case Selection.Interior.Color
        Case rgbLightGrey: Selection.Interior.Color = vbRed
            If Not Selection.Value = MineSign Then
                lMarkedFiledWOMine = lMarkedFiledWOMine + 1
            End If
        Case vbRed: Selection.Interior.Color = rgbLightGrey
            If Not Selection.Value = MineSign Then
                lMarkedFiledWOMine = lMarkedFiledWOMine - 1
            End If
        Case Else
        End Select
    ElseIf Selection.Interior.Color = rgbLightGrey Then
        Call UnHide(Selection, True)
    End If
    
    If Not bLost And Not bWon And lUnhidenFieldsCount + NumberOfMines = 16 * 32 And lMarkedFiledWOMine = 0 Then
        bWon = True
        MsgBox "(^.^)"
    End If
    
    bRightClicked = False
    bSelectionChanged = False
End Sub
