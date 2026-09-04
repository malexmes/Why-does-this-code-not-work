Option Explicit

' The Econometrics Game v3. All game state lives on the State sheet.
' Buttons are cells. Each app sheet passes clicks here from Worksheet_SelectionChange.

Private Const MAXSTEP As Long = 21

Private Function St() As Worksheet
    Set St = ThisWorkbook.Worksheets("State")
End Function

Private Function Stp() As Worksheet
    Set Stp = ThisWorkbook.Worksheets("Steps")
End Function

Private Function CurStep() As Long
    CurStep = CLng(Val(St.Range("B1").Value))
End Function

Private Function StepType() As String
    StepType = CStr(Stp.Cells(CurStep + 1, 6).Value)
End Function

Public Sub HandleClick(ByVal ws As Worksheet, ByVal Target As Range)
    Dim a As String
    If Target.Cells.Count > 40 Then Exit Sub
    a = Target.Cells(1, 1).Address(False, False)
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo fin
    Dispatch ws.Name, a
fin:
    On Error Resume Next
    ActiveSheet.Range("A1").Select
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub Dispatch(ByVal sheetName As String, ByVal a As String)
    Select Case sheetName
        Case "Play": PlayClick a
        Case "Budget": BudgetClick a
        Case "Long game": ReadClick a, 3
        Case "Giving": ReadClick a, 4
        Case "Finish": FinishClick a
    End Select
End Sub

Private Sub PlayClick(ByVal a As String)
    Select Case a
        Case "B29": GoBack
        Case "T29": GoNext
        Case "G29": St.Range("B15").Value = 1 - Val(St.Range("B15").Value)
        Case "L29": St.Range("B14").Value = 1 - Val(St.Range("B14").Value)
        Case "Q29": ResetGame
        Case "U16": NudgeDial -1, True
        Case "V16": NudgeDial -1, False
        Case "X16": NudgeDial 1, False
        Case "Y16": NudgeDial 1, True
        Case "T11": ChooseOption 1
        Case "T13": ChooseOption 2
        Case "T15": ChooseOption 3
        Case "U11": NudgeChannel 1, -1
        Case "U12": NudgeChannel 2, -1
        Case "U13": NudgeChannel 3, -1
        Case "U14": NudgeChannel 4, -1
        Case "U15": NudgeChannel 5, -1
        Case "V11": NudgeChannel 1, 1
        Case "V12": NudgeChannel 2, 1
        Case "V13": NudgeChannel 3, 1
        Case "V14": NudgeChannel 4, 1
        Case "V15": NudgeChannel 5, 1
    End Select
End Sub

Private Sub GoNext()
    Dim s As Long
    s = CurStep
    If s < MAXSTEP Then
        SetStep s + 1
    Else
        ShowStage 2
    End If
End Sub

Private Sub GoBack()
    Dim s As Long
    s = CurStep
    If s > 1 Then SetStep s - 1
End Sub

Private Sub SetStep(ByVal n As Long)
    If n < 1 Then n = 1
    If n > MAXSTEP Then n = MAXSTEP
    St.Range("B1").Value = n
    St.Range("B14").Value = 0
    St.Range("B15").Value = 0
End Sub

Public Sub ShowStage(ByVal n As Long)
    Dim nm As String
    Select Case n
        Case 1: nm = "Play"
        Case 2: nm = "Budget"
        Case 3: nm = "Long game"
        Case 4: nm = "Giving"
        Case Else: nm = "Finish": n = 5
    End Select
    St.Range("B30").Value = n
    ThisWorkbook.Worksheets(nm).Activate
    On Error Resume Next
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayGridlines = False
    ActiveSheet.Range("A1").Select
End Sub

Private Sub NudgeDial(ByVal sign As Long, ByVal big As Boolean)
    Dim r As Long, dialId As Long, stepSize As Double, v As Double
    If StepType <> "dial" Then Exit Sub
    r = CurStep + 1
    dialId = CLng(Val(Stp.Cells(r, 7).Value))
    If dialId = 0 Then Exit Sub
    If big Then
        stepSize = Val(Stp.Cells(r, 9).Value)
    Else
        stepSize = Val(Stp.Cells(r, 8).Value)
    End If
    v = Val(St.Cells(1 + dialId, 2).Value) + sign * stepSize
    If v < 0 Then v = 0
    If dialId = 5 And v > 90 Then v = 90
    St.Cells(1 + dialId, 2).Value = Round(v, 2)
End Sub

Private Sub NudgeChannel(ByVal k As Long, ByVal sign As Long)
    Dim v As Double
    If StepType <> "dials5" Then Exit Sub
    v = Val(St.Cells(7 + k, 2).Value) + sign * 1
    If v < 0 Then v = 0
    St.Cells(7 + k, 2).Value = v
End Sub

Private Sub ChooseOption(ByVal opt As Long)
    Dim q As Long
    If StepType <> "quiz" Then Exit Sub
    q = CLng(Val(Stp.Cells(CurStep + 1, 19).Value))
    If q = 0 Then Exit Sub
    St.Cells(16 + q, 2).Value = opt
End Sub

Public Sub ResetGame()
    Dim i As Long
    For i = 2 To 12
        St.Cells(i, 2).Value = 0
    Next i
    St.Range("B14").Value = 0
    St.Range("B15").Value = 0
    For i = 17 To 26
        St.Cells(i, 2).Value = 0
    Next i
    For i = 51 To 63
        St.Cells(i, 2).Value = St.Cells(i, 3).Value
    Next i
    St.Range("B64").Value = 0
    St.Range("B1").Value = 1
    ShowStage 1
End Sub

' ---- Budget sheet
Private Sub BudgetClick(ByVal a As String)
    Dim col As String, rowNum As Long
    Select Case a
        Case "B29": ShowStage 1
        Case "T29": ShowStage 3
        Case "G29": St.Range("B64").Value = 1 - Val(St.Range("B64").Value)
        Case "L29": BudgetReset
        Case "T20": St.Range("B24").Value = 1
        Case "T21": St.Range("B24").Value = 2
        Case "T22": St.Range("B24").Value = 3
        Case Else
            If Len(a) >= 3 Then
                col = Left(a, 1)
                rowNum = CLng(Val(Mid(a, 2)))
                If rowNum >= 11 And rowNum <= 23 Then
                    If col = "M" Then BudgetNudge rowNum - 10, -1
                    If col = "N" Then BudgetNudge rowNum - 10, 1
                End If
            End If
    End Select
End Sub

Private Sub BudgetNudge(ByVal i As Long, ByVal sign As Long)
    Dim baseSpend As Double, v As Double, stepSize As Double
    If Val(St.Range("B64").Value) = 1 Then St.Range("B64").Value = 0
    baseSpend = Val(St.Cells(50 + i, 3).Value)
    stepSize = Round(baseSpend * 0.1, 0)
    If stepSize < 1 Then stepSize = 1
    v = Val(St.Cells(50 + i, 2).Value) + sign * stepSize
    If v < baseSpend * 0.5 Then v = Round(baseSpend * 0.5, 0)
    If v > baseSpend * 1.5 Then v = Round(baseSpend * 1.5, 0)
    St.Cells(50 + i, 2).Value = v
End Sub

Private Sub BudgetReset()
    Dim i As Long
    For i = 51 To 63
        St.Cells(i, 2).Value = St.Cells(i, 3).Value
    Next i
    St.Range("B64").Value = 0
End Sub

' ---- Reading chapters (Long game = stage 3, Giving = stage 4)
Private Sub ReadClick(ByVal a As String, ByVal stage As Long)
    Dim q As Long
    q = 6 + stage
    Select Case a
        Case "B29": ShowStage stage - 1
        Case "T29": ShowStage stage + 1
        Case "T20": St.Cells(16 + q, 2).Value = 1
        Case "T21": St.Cells(16 + q, 2).Value = 2
        Case "T22": St.Cells(16 + q, 2).Value = 3
    End Select
End Sub

Private Sub FinishClick(ByVal a As String)
    Select Case a
        Case "B29": ShowStage 4
        Case "T29": ResetGame
    End Select
End Sub

Public Sub OpenGame()
    Dim n As Long
    n = CLng(Val(St.Range("B30").Value))
    If n < 1 Then n = 1
    ShowStage n
End Sub
