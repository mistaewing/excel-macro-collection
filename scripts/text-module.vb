'Converts the text in the selected range to upper case
Sub ConvertToUpperCase()
    Dim Rng As Range
    For Each Rng In Selection
        If Application.WorksheetFunction.IsText(Rng) Then
            Rng.Value = UCase(Rng)
        End If
    Next
End Sub

'Converts the text in the selected range to lower case
Sub ConvertToLowerCase()
    Dim Rng As Range
    For Each Rng In Selection
        If Application.WorksheetFunction.IsText(Rng) Then
            Rng.Value= LCase(Rng)
        End If
    Next
End Sub

'Converts the text in the selected range to title case where the first letter is upper case the rest is lower
Sub ConvertToTitleCase()
    Dim Rng As Range
    For Each Rng In Selection
        If WorksheetFunction.IsText(Rng) Then
            Rng.Value= WorksheetFunction.Proper(Rng.Value)
        End If
    Next
End Sub

'Converts the text in the selected range to sentence case
Sub ConvertToSentenceCase()
    Dim Rng As Range
    For Each Rng In Selection
        If WorksheetFunction.IsText(Rng) Then
            Rng.Value= UCase(Left(Rng, 1)) & LCase(Right(Rng, Len(Rng) -1))
        End If
    Next rng
End Sub

'Writes a sequence of numbers in a column from the active cells.
Sub AddSerialNumbers()
    Dim i As Integer
    On Error GoTo Last
    i = InputBox("Enter Value", "Enter Serial Numbers")
    For i = 1 To i
        ActiveCell.Value = i
        ActiveCell.Offset(1, 0).Activate
    Next i
    Last:Exit Sub
End Sub