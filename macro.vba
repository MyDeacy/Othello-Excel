Option Explicit

Dim turn As Integer
Dim game As Boolean
Dim ch1(8) As Integer
Dim ch2(8) As Integer

Sub reset()
    turn = 0
    ch1(0) = 1
    ch1(1) = 1
    ch1(2) = 0
    ch1(3) = -1
    ch1(4) = -1
    ch1(5) = -1
    ch1(6) = 0
    ch1(7) = 1
    
    ch2(0) = 0
    ch2(1) = 1
    ch2(2) = 1
    ch2(3) = 1
    ch2(4) = 0
    ch2(5) = -1
    ch2(6) = -1
    ch2(7) = -1
    ActiveSheet.Unprotect
    Range("A1:H8").Value = ""
    Cells(4, 4).Value = "●"
    Cells(4, 5).Value = "〇"
    Cells(5, 5).Value = "●"
    Cells(5, 4).Value = "〇"
    game = True
    bumpTurn
    ActiveSheet.Protect
End Sub


Public Function check(ByVal cell As Object)
    Dim column As Integer
    Dim row As Integer
    Dim flag As Integer
    Dim row2 As Integer
    Dim column2 As Integer
    Dim my As String
    Dim yo As String
    Dim i As Integer
    Dim boo As Boolean
    On Error GoTo Exception
    
    column = cell.column
    row = cell.row
    
    If Not cell.Value = "" Or Not (column <= 8 And row <= 8 And column >= 1 And row >= 1) Then
        Exit Function
    End If
    
    ActiveSheet.Unprotect
    
    flag = 0
    boo = False
    
    my = IIf(turn = 1, "●", "〇")
    yo = IIf(turn = 1, "〇", "●")

    i = 0
    While i < 8
        row = cell.row
        column = cell.column
        If column + ch2(i) <= 8 And (row + ch1(i)) <= 8 And (column + ch2(i)) >= 1 And (row + ch1(i)) >= 1 Then
            flag = IIf(column + ch2(i) <= 8 And (row + ch1(i)) <= 8 And (column + ch2(i)) >= 1 And (row + ch1(i)) >= 1 _
            , IIf(Cells(row + ch1(i), column + ch2(i)).Value = yo, 1, 0), 0)
        End If
            
            
        Do While (column + ch2(i)) <= 8 And (row + ch1(i)) <= 8 And (column + ch2(i)) >= 1 And (row + ch1(i)) >= 1
            row = row + ch1(i)
            column = column + ch2(i)
            If Cells(row, column).Value = "" Then
                Exit Do
            End If
                
            If Cells(row, column).Value = my And flag = 1 Then
                    row2 = cell.row
                    column2 = cell.column
                    flag = 0
                    While IIf(ch1(i) = 0, True, Not row = row2) And IIf(ch2(i) = 0, True, Not column = column2)
                        Cells(row2, column2).Value = my
                        row2 = row2 + ch1(i)
                        column2 = column2 + ch2(i)
                    Wend
                    boo = True
                    Exit Do
            End If
        Loop
        i = i + 1
    Wend
    
    If Not boo Then
        cell.Value = ""
        ActiveSheet.Protect
        Exit Function
    End If
    bumpTurn
    ActiveSheet.Protect
    
Exit Function

Exception:
    ActiveSheet.Protect
    MsgBox "判定ができませんでした。" & vbCrLf & "正確に指定してください。", vbExclamation
End Function



Private Sub Worksheet_BeforeRightClick(ByVal target As Range, Cancel As Boolean)
    Cancel = True
    If game Then
        Dim mark As String
            ActiveSheet.Unprotect
            check target
    End If

End Sub

Sub pass()
    bumpTurn
End Sub

Function bumpTurn()
    Dim cell As Range
    ActiveSheet.Unprotect
    Range("A1:H8").Interior.ColorIndex = 50

    For Each cell In Range("A1:H8")
        markTarget cell
    Next
    turn = IIf(turn = 0, 1, 0)
    Range("A1:H8").Borders.ColorIndex = 1
    Cells(5, 10).Value = IIf(turn = 1, "●", "〇")

    ActiveSheet.Protect
End Function


Sub countBW()
    Dim countb As Integer
    Dim countw As Integer
    countb = WorksheetFunction.CountIf(Range("A1:H8"), "●")
    countw = WorksheetFunction.CountIf(Range("A1:H8"), "〇")
    
    MsgBox "● x " & countb & " : " & "〇 x " & countw

End Sub


Function markTarget(ByVal cell As Range)
    ActiveSheet.Unprotect
    Dim column As Integer
    Dim row As Integer
    Dim flag As Integer
    Dim my As String
    Dim yo As String
    Dim i As Integer
    column = cell.column
    row = cell.row
    flag = 0
    my = IIf(turn = 0, "●", "〇")
    yo = IIf(turn = 0, "〇", "●")
    If Not cell.Value = "" Then
        Exit Function
    End If
    
    i = 0
    While i < 8
        row = cell.row
        column = cell.column
        If column + ch2(i) <= 8 And (row + ch1(i)) <= 8 And (column + ch2(i)) >= 1 And (row + ch1(i)) >= 1 Then
            flag = IIf(column + ch2(i) <= 8 And (row + ch1(i)) <= 8 And (column + ch2(i)) >= 1 And (row + ch1(i)) >= 1 _
            , IIf(Cells(row + ch1(i), column + ch2(i)).Value = yo, 1, 0), 0)
        End If
            
        Do While (column + ch2(i)) <= 8 And (row + ch1(i)) <= 8 And (column + ch2(i)) >= 1 And (row + ch1(i)) >= 1
            row = row + ch1(i)
            column = column + ch2(i)
            If Cells(row, column).Value = "" Then
                Exit Do
            End If
            If Cells(row, column).Value = my And flag = 1 Then
                cell.Interior.ColorIndex = 4
                Exit Function
            End If
        Loop
        i = i + 1
    Wend

End Function

