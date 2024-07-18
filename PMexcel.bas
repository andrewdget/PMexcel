Attribute VB_Name = "PMexcel"
Sub run_PMexcel()

lock_cells False

Call format_header
Call dropdown_validation
Call remove_blank_rows
Call format_rows

lock_cells True

End Sub

Private Sub format_header()

With ActiveSheet
    .Columns("A").ColumnWidth = 10
    .Columns("B").ColumnWidth = 60
    .Columns("C:E").ColumnWidth = 10
    .Columns("F").ColumnWidth = 20
    .Columns("G:Q").ColumnWidth = 10
End With

Range("A1").Value = "Activity ID"
Range("A1:A2").Merge
Range("AA:AA").NumberFormat = "@"

Range("B1").Value = "Activity Description"
Range("B1:B2").Merge

Range("C1").Value = "Planned Duration"
With Range("C1:C2")
    .Merge
    .WrapText = True
End With

Range("D1").Value = "Calculated Duration"
With Range("D1:D2")
    .Merge
    .WrapText = True
End With

Range("E1").Value = "Float"
Range("E1:E2").Merge

Range("F1").Value = "Predecessor(s)"
Range("F1:F2").Merge

Range("G1").Value = "Need to"
Range("G1:H1").Merge

Range("I1").Value = "Forcast to"
Range("I1:J1").Merge

Range("K1").Value = "Plan to"
Range("K1:L1").Merge

Range("M1").Value = "Actual"
Range("M1:N1").Merge

Range("G2, I2, K2, M2").Value = "Start"
Range("H2, J2, L2, N2").Value = "Finish"

Range("O1").Value = "Schedule Type"
With Range("O1:O2")
    .Merge
    .WrapText = True
End With

Range("P1").Value = "Task Type"
Range("P1:P2").Merge

Range("Q1").Value = "Task Status"
With Range("Q1:Q2")
    .Merge
    .WrapText = True
End With

With Range("A1:Q2")
    .Font.Bold = True
    .Font.Color = vbWhite
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .Interior.Color = RGB(21, 96, 130)
    .Borders.LineStyle = xlContinuous
End With

End Sub

Private Sub dropdown_validation()

Dim lastrow
lastrow = ActiveSheet.UsedRange.Rows.Count

Dim validation_list
validation_list = "Calendar, Workdays"
Range("O3:O" & lastrow).Validation.Delete
With Range("O3:O" & lastrow).Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=validation_list
    .IgnoreBlank = True
    .InCellDropdown = True
    .ShowInput = True
    .ShowError = True
End With
    

End Sub

Private Sub format_rows()

Set wf = WorksheetFunction

Dim lastrow
Dim numlayers
Dim id
Dim layer
Dim maxcolor
Dim colorstep
Dim layercolor

lastrow = ActiveSheet.UsedRange.Rows.Count
numlayers = 0
For Row = 3 To lastrow
    id = Range("A" & Row).Value
    layer = Len(id) - Len(wf.Substitute(id, ".", "")) + 1
    If layer > numlayers Then numlayers = layer
    
    If layer > 1 Then Range("B" & Row).IndentLevel = layer - 1
    Range("A" & Row & ":Q" & Row).Borders.LineStyle = xlContinuous
Next Row

maxcolor = 140
If numlayers > 1 Then
    colorstep = wf.RoundDown((255 - maxcolor) / (numlayers - 1), 0)
    For Row = 3 To lastrow
        If Not wf.CountIf(Range("A3:A" & lastrow), Range("A" & Row).Value & "*") = 1 Then
            id = Range("A" & Row).Value
            layer = Len(id) - Len(wf.Substitute(id, ".", "")) + 1
            If Not layer = numlayer Then
                layercolor = maxcolor + (colorstep * (layer - 1))
                Range("A" & Row & ":" & "Q" & Row).Interior.Color = RGB(layercolor, layercolor, layercolor)
            Else
                 Range("A" & Row & ":" & "Q" & Row).Interior.Color = xlNone
            End If
        Else
             Range("A" & Row & ":" & "Q" & Row).Interior.Color = xlNone
        End If
    Next Row
End If

End Sub

Private Sub remove_blank_rows()

Dim lastrow

lastrow = ActiveSheet.UsedRange.Rows.Count
For Row = lastrow To 3 Step -1
    If IsEmpty(Range("A" & Row).Value) And IsEmpty(Range("B" & Row).Value) Then Range("A" & Row).EntireRow.Delete
Next Row

End Sub

Private Sub lock_cells(state As Boolean)

Dim lastrow
lastrow = ActiveSheet.UsedRange.Rows.Count

ActiveSheet.Unprotect

Range("A1:Q2").Locked = state
Range("D3:E" & lastrow).Locked = state
Range("G3:J" & lastrow).Locked = state
Range("P3:Q" & lastrow).Locked = state

If state Then
    ActiveSheet.Protect
Else
    ActiveSheet.Unprotect
End If

End Sub
