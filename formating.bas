Attribute VB_Name = "formating"
Sub subproject_formating()

Set WF = WorksheetFunction

Dim LastRow
Dim NumLayers
Dim ID
Dim Layer
Dim MaxColor
Dim ColorStep
Dim LayerColor

LastRow = ActiveSheet.UsedRange.Rows.Count
NumLayers = 0
For Row = 3 To LastRow
    ID = Range("A" & Row).Value
    Layer = Len(ID) - Len(WF.Substitute(ID, ".", "")) + 1
    If Layer > numlayer Then numlayer = Layer
    
    Range("B" & Row).IndentLevel = Layer - 1
    Range("A" & Row & ":" & "M" & Row).Borders.LineStyle = xlContinuous
Next Row

MaxColor = 140
ColorStep = WF.RoundDown((255 - MaxColor) / (numlayer - 1), 0)
For Row = 3 To LastRow
    If Not WF.CountIf(Range("A3:A" & LastRow), Range("A" & Row).Value & "*") = 1 Then
        ID = Range("A" & Row).Value
        Layer = Len(ID) - Len(WF.Substitute(ID, ".", "")) + 1
        If Not Layer = numlayer Then
            LayerColor = MaxColor + (ColorStep * (Layer - 1))
            Range("A" & Row & ":" & "M" & Row).Interior.Color = RGB(LayerColor, LayerColor, LayerColor)
        Else
             Range("A" & Row & ":" & "M" & Row).Interior.Color = xlNone
        End If
    Else
         Range("A" & Row & ":" & "M" & Row).Interior.Color = xlNone
    End If
Next Row

End Sub
Sub reset_formating()

Dim LastRow

LastRow = ActiveSheet.UsedRange.Rows.Count
ActiveSheet.Range("A1:M" & LastRow).ClearFormats

ActiveSheet.Range("A1:M2").ClearContents

With ActiveSheet
    .Columns("A").ColumnWidth = 9
    .Columns("B").ColumnWidth = 60
    .Columns("C").ColumnWidth = 8
    .Columns("D").ColumnWidth = 20
    .Columns("E:M").ColumnWidth = 10
End With

Range("A1").Value = "Activity ID"
Range("A1:A2").Merge
Range("A2:A" & LastRow).NumberFormat = "@"

Range("B1").Value = "Activity Description"
Range("B1:B2").Merge

Range("C1").Value = "Duration"
Range("C1:C2").Merge

Range("D1").Value = "Predecessor(s)"
Range("D1:D2").Merge

Range("E1").Value = "Need to"
Range("E1:F1").Merge

Range("G1").Value = "Plan to"
Range("G1:H1").Merge

Range("I1").Value = "Actual"
Range("I1:J1").Merge

Range("E2, G2, I2").Value = "Start"
Range("F2, H2, J2").Value = "Finish"

Range("K1").Value = "Task Type"
Range("K1:K2").Merge

Range("L1").Value = "Calendar Type"
With Range("L1:L2")
    .Merge
    .WrapText = True
End With

Range("M1").Value = "Schedule Type"
With Range("M1:M2")
    .Merge
    .WrapText = True
End With

With Range("A1:M2")
    .Font.Bold = True
    .Font.Color = vbWhite
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .Interior.Color = RGB(21, 96, 130)
    .Borders.LineStyle = xlContinuous
End With

For Row = 3 To LastRow
    If IsEmpty(Range("A" & Row).Value) And IsEmpty(Range("B" & Row).Value) Then Range("A" & Row).EntireRow.Delete
Next Row

Call subproject_formating

Range("B1").Value = "Activity Description"

End Sub
