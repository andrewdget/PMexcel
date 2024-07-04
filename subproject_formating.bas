Attribute VB_Name = "Module1"
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
    Range("A" & Row & ":" & "L" & Row).Borders.LineStyle = xlContinuous
Next Row

MaxColor = 140
ColorStep = WF.RoundDown((255 - MaxColor) / (numlayer - 1), 0)
For Row = 3 To LastRow
    If Not WF.CountIf(Range("A3:A" & LastRow), Range("A" & Row).Value & "*") = 1 Then
        ID = Range("A" & Row).Value
        Layer = Len(ID) - Len(WF.Substitute(ID, ".", "")) + 1
        If Not Layer = numlayer Then
            LayerColor = MaxColor + (ColorStep * (Layer - 1))
            Range("A" & Row & ":" & "L" & Row).Interior.Color = RGB(LayerColor, LayerColor, LayerColor)
        Else
             Range("A" & Row & ":" & "L" & Row).Interior.Color = xlNone
        End If
    Else
         Range("A" & Row & ":" & "L" & Row).Interior.Color = xlNone
    End If
Next Row

End Sub
