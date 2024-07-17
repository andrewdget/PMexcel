Attribute VB_Name = "PMexcel"
Sub project_setup()

Dim lastrow

lastrow = ActiveSheet.UsedRange.Rows.Count
ActiveSheet.Range("A1:P" & lastrow).ClearFormats
ActiveSheet.Range("A1:P" & lastrow).ClearContents

With ActiveSheet
    .Columns("A").ColumnWidth = 10
    .Columns("B").ColumnWidth = 60
    .Columns("C:E").ColumnWidth = 10
    .Columns("F").ColumnWidth = 20
    .Columns("G:P").ColumnWidth = 10
End With

Range("A1").Value = "Activity ID"
Range("A1:A2").Merge
Range("A2:A" & lastrow).NumberFormat = "@"

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

Range("O1").Value = "Task Type"
Range("O1:O2").Merge

Range("P1").Value = "Calendar Type"
With Range("P1:P2")
    .Merge
    .WrapText = True
End With

With Range("A1:P2")
    .Font.Bold = True
    .Font.Color = vbWhite
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .Interior.Color = RGB(21, 96, 130)
    .Borders.LineStyle = xlContinuous
End With

End Sub

Sub refresh()

Call remove_blanks
Call general_formating

End Sub

Sub general_formating()

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
    Range("A" & Row & ":P" & Row).Borders.LineStyle = xlContinuous
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
                Range("A" & Row & ":" & "P" & Row).Interior.Color = RGB(layercolor, layercolor, layercolor)
            Else
                 Range("A" & Row & ":" & "P" & Row).Interior.Color = xlNone
            End If
        Else
             Range("A" & Row & ":" & "P" & Row).Interior.Color = xlNone
        End If
    Next Row
End If

End Sub

Sub remove_blanks()

Dim lastrow

lastrow = ActiveSheet.UsedRange.Rows.Count
For Row = lastrow To 3 Step -1
    If IsEmpty(Range("A" & Row).Value) And IsEmpty(Range("B" & Row).Value) Then Range("A" & Row).EntireRow.Delete
Next Row

End Sub
