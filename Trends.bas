Attribute VB_Name = "Trends"
'Sub FormatTrendSheet()
'With ActiveSheet
'    .Range("A:A").ColumnWidth = 3
'    .Range("A:A").RowHeight = 15
'    '.Range("B:B").ColumnWidth = 15
'    .Range("B:K").ColumnWidth = 11
'    .Range("L:L").ColumnWidth = 3
'    With .Range("A1:L1")
'        .MergeCells = True
'        .Interior.color = 14277081
'        .RowHeight = 40
'    End With
'    With .Range("A2:L2")
'        .MergeCells = True
'        .Interior.color = 6968388
'        .RowHeight = 4
'    End With
'    With .Range("A3:L3")
'        .MergeCells = True
'        .Font.color = -9808828
'        .RowHeight = 32
'        .HorizontalAlignment = xlCenter
'        .Font.Name = "Calibri"
'        .Font.Size = 18
'    End With
'    With .Range("A4:L4")
'        .MergeCells = True
'        .Font.color = -9808828
'        .Font.Bold = True
'        .RowHeight = 25
'        .HorizontalAlignment = xlCenter
'        .Font.Name = "Calibri"
'        .Font.Size = 18
'    End With
'    With .Range("A5:L5")
'        .MergeCells = True
'        .HorizontalAlignment = xlCenter
'    End With
'    With .Range("B6:K6")
'        .MergeCells = True
'        .Interior.color = 6968388
'        .HorizontalAlignment = xlLeft
'        .Font.color = RGB(255, 255, 255)
'        .Font.Bold = True
'        .Font.Name = "Calibri"
'        .Font.Size = 12
'    End With
'
'End With
'
'End Sub

Sub FormatChartBlock(chartheadrow, trendname, x)
If x = 0 Then
    With ActiveSheet.Range("A" & chartheadrow & ":J" & chartheadrow)
        .MergeCells = True
        .Interior.color = 8421504
        .HorizontalAlignment = xlLeft
        '.Font.color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Value = trendname
    End With
    With ActiveSheet.Range("A" & chartheadrow + 1 & ":J" & chartheadrow + 12)
        .MergeCells = True
        '.Border = True
    End With
End If
If x = 48 Then
    With ActiveSheet.Range("B" & chartheadrow & ":K" & chartheadrow)
        .MergeCells = True
        .Interior.color = 10921638
        .HorizontalAlignment = xlLeft
        '.Font.color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Value = trendname
    End With
    With ActiveSheet.Range("B" & chartheadrow + 1 & ":K" & chartheadrow + 12)
        .MergeCells = True
        '.Border = True
    End With
End If
If x = 96 Then
    With ActiveSheet.Range("C" & chartheadrow & ":L" & chartheadrow)
        .MergeCells = True
        .Interior.color = 14277081
        .HorizontalAlignment = xlLeft
        '.Font.color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Value = trendname
    End With
    With ActiveSheet.Range("C" & chartheadrow + 1 & ":L" & chartheadrow + 12)
        .MergeCells = True
        '.Border = True
    End With
End If
End Sub


Sub CreateTrends()
Application.ScreenUpdating = False
Dim usetable As ListObject
Set usetable = Sheets("TrendData").ListObjects("TrendData")
chartheadrow = 1

'y = 15.5

For c = 1 To usetable.ListRows.count
    With Sheets("Trends")
        trendname = Sheets("TrendData").Range("A" & c + 1).Value
        
        If Left(trendname, 1) = "C" Then x = 0 'And coderow = c
        If Left(trendname, 1) = "A" Then x = 48 'And arearow = c
        If Left(trendname, 1) = "F" Then x = 96 'And facilityrow = c
        
        Call FormatChartBlock(chartheadrow, trendname, x)
        y = Cells(chartheadrow + 1, 1).Top
        chartheadrow = chartheadrow + 14
        'ADD CHART
        .Shapes.AddChart(xlLineMarkers, x, y, 480, 180).Select
        With ActiveChart
            
            .SetSourceData Source:=usetable.ListRows(c).Range, PlotBy:=xlRows
            .SeriesCollection(1).XValues = usetable.HeaderRowRange.Offset(0, 1) ' + 1 'Range.Offset(0, 1) '("='TrendData'!$B$1:$C$1" ' & lastcolumn
            .HasLegend = False
            .HasTitle = False
            .SeriesCollection(1).MarkerStyle = 8
            .SeriesCollection(1).MarkerSize = 7
            .SeriesCollection(1).Format.Line.Weight = 2
            .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(68, 84, 106)
            .SeriesCollection(1).MarkerBackgroundColor = RGB(255, 255, 255)
            .Axes(xlValue).HasTitle = True
            'Set cmt = Sheets("TrendData").Range("A" & c + 1).Comment.Text
            .Axes(xlValue).AxisTitle.Caption = "Quantity - " & Sheets("TrendData").Range("A" & c + 1).Comment.Text
            .HasTitle = False
            '.Axes(xlValue).AxisTitle.Font.Name = "bookman"
            '.Axes(xlValue).AxisTitle.Font.Size = 10
            '.Axes(xlValue).AxisTitle.Characters(9, 5).Font.Italic = True
        End With
        'Setup Chart block format
        'trendname = Sheets("TrendData").Range("A" & c + 1).Value
        'Call FormatChartBlock(chartheadrow, trendname)
        'y = y + 210.5
        'chartheadrow = chartheadrow + 14
    End With
Next

Application.ScreenUpdating = True

End Sub


Sub XYCoordinates()
With Selection
    x = .Left
End With
Debug.Print "X = " & x & vbNewLine & "Y = " & y
End Sub


Sub hyperlinksshortcut()
x2 = 1
For x = 4 To 141 'NEED TO CHANGE HARD-CODED VALUES TO LAST ROW
    Sheets("Comparison").hyperlinks.Add Anchor:=Sheets("Comparison").Range("F" & x), Address:="", SubAddress:="'Trends'!A" & x2
    Sheets("Trends").hyperlinks.Add Anchor:=Sheets("Trends").Range("A" & x2), Address:="", SubAddress:="'Comparison'!F" & x
    Sheets("Trends").hyperlinks.Add Anchor:=Sheets("Trends").Range("b" & x2), Address:="", SubAddress:="'Comparison'!F" & x
    Sheets("Trends").hyperlinks.Add Anchor:=Sheets("Trends").Range("c" & x2), Address:="", SubAddress:="'Comparison'!F" & x
    x2 = x2 + 14
Next
End Sub

