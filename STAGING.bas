Attribute VB_Name = "STAGING"
'Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'RUN TREND REPORT ON THE EXISTING COST CODE COMPARISON SHEET
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub TrendGraphs()
    Application.ScreenUpdating = False

    'DECLARE AND SET WORKBOOKS AND WORKSHEETS
    Dim wb_master As Workbook
    Dim ws_compare As Worksheet
    Dim ws_trenddata As Worksheet
    Dim ws_trends As Worksheet

    Set wb_master = ActiveWorkbook
    Set ws_compare = Worksheets("Comparison")
    Set ws_trenddata = Worksheets("TrendData")
    Set ws_trends = Worksheets("Trends")

'DECLARE AND SET COLUMNS AND ROWS
    Dim firstrow As Integer
    Dim lastrow As Long
    Dim lasttrendrow As Long
    Dim itemrow As Long
    Dim newtrendrow As Long
    Dim trenddatacolumn As Integer

    
    
    firstrow = Application.Match("Item", ws_compare.Range("A1:A10"), False) + 1
    lastrow = ws_compare.Cells(Rows.count, 1).End(xlUp).Row
    lasttrendrow = ws_trenddata.Cells(Rows.count, 1).End(xlUp).Row
    newtrendrow = lasttrendrow + 1
    trenddatacolumn = ws_trenddata.Cells(1, Columns.count).End(xlToLeft).Column + 1
    

'DECLARE AND SET OTHERS'
    Dim item As Variant
    Dim costcode As String
    Dim area As String
    Dim facility As String
    Dim modelqty As Variant
    Dim estimateqty As Variant
    Dim uom As String
    Dim trendfound As Variant
    Dim tcfound As Object
    Dim setfound As Variant
    
    ws_trenddata.Cells(1, trenddatacolumn) = Date
    
    
    'IF TRENDS SHEET EXISTS, DELETE IT AND RE-CREATE IT
    For Each sheet In ActiveWorkbook.Sheets
        If sheet.Name = "Trends" Then
            Application.DisplayAlerts = False
            Worksheets("Trends").Delete
            Application.DisplayAlerts = True
            Worksheets.Add(After:=Sheets("Comparison")).Name = "Trends"
            Set ws_trends = Worksheets("Trends")
            ws_trends.Range("A1").Value = "EVLLRT - Trend Charts"
                
        End If
    Next
            
    ws_trends.Activate


With ws_compare
    For Each item In .Range(Cells(firstrow, 1).Address, Cells(lastrow, 1).Address)
        itemrow = item.Row
        
        'NEED TO CHANGE STATIC COLUMN REFERENCE
        modelqty = .Cells(itemrow, 6).Value
        estimateqty = .Cells(itemrow, 7).Value
        uom = .Cells(itemrow, 4).Value
        
        If modelqty = "No QTY" Then modelqty = 0
        If estimateqty = "No Estimate" Then estimateqty = 0
        
        If item.IndentLevel = 0 Then
            costcode = item
            trendfound = Application.Match(costcode, ws_trenddata.Range("A1:A" & lasttrendrow), False)
            If IsError(trendfound) Then
                With ws_trenddata
                    .Cells(newtrendrow, 1).Value = costcode
                    .Cells(newtrendrow, 1).ClearComments
                    .Cells(newtrendrow, 1).AddComment
                    .Cells(newtrendrow, 1).Comment.Text Text:=uom
                    .Range(Cells(newtrendrow, 1).Address, Cells(newtrendrow + 1, 1).Address).Merge
                    .Cells(newtrendrow, trenddatacolumn).Value = modelqty
                    .Cells(newtrendrow + 1, trenddatacolumn).Value = estimateqty
                    newtrendrow = newtrendrow + 2
                End With
            Else
                With ws_trenddata
                    .Cells(trendfound, trenddatacolumn).Value = modelqty
                    .Cells(trendfound + 1, trenddatacolumn).Value = estimateqty
                End With
            End If
        End If
        If item.IndentLevel = 1 Then
            area = item
            trendfound = Application.Match(costcode & " | " & area, ws_trenddata.Range("A1:A" & lasttrendrow), False)
            If IsError(trendfound) Then
                With ws_trenddata
                    .Cells(newtrendrow, 1).Value = costcode & " | " & area
                    .Cells(newtrendrow, 1).ClearComments
                    .Cells(newtrendrow, 1).AddComment
                    .Cells(newtrendrow, 1).Comment.Text Text:=uom
                    .Range(Cells(newtrendrow, 1).Address, Cells(newtrendrow + 1, 1).Address).Merge
                    .Cells(newtrendrow, trenddatacolumn).Value = modelqty
                    .Cells(newtrendrow + 1, trenddatacolumn).Value = estimateqty
                    newtrendrow = newtrendrow + 2
                End With
            Else
                With ws_trenddata
                    .Cells(trendfound, trenddatacolumn).Value = modelqty
                    .Cells(trendfound + 1, trenddatacolumn).Value = estimateqty
                End With
            End If
        End If
        If item.IndentLevel = 3 Then
            facility = item
            trendfound = Application.Match(costcode & " | " & area & " | " & facility, ws_trenddata.Range("A1:A" & lasttrendrow), False)
            If IsError(trendfound) Then
                With ws_trenddata
                    .Cells(newtrendrow, 1).Value = costcode & " | " & area & " | " & facility
                    .Cells(newtrendrow, 1).ClearComments
                    .Cells(newtrendrow, 1).AddComment
                    .Cells(newtrendrow, 1).Comment.Text Text:=uom
                    .Range(Cells(newtrendrow, 1).Address, Cells(newtrendrow + 1, 1).Address).Merge
                    .Cells(newtrendrow, trenddatacolumn).Value = modelqty
                    .Cells(newtrendrow + 1, trenddatacolumn).Value = estimateqty
                    newtrendrow = newtrendrow + 2
                End With
            Else
                With ws_trenddata
                    .Cells(trendfound, trenddatacolumn).Value = modelqty
                    .Cells(trendfound + 1, trenddatacolumn).Value = estimateqty
                End With
            End If
        End If
        
        
    Next
End With
    
    
'CLEAR EXISTING TREND CHARTS AND CREATE NEW ONES
ws_trends.UsedRange.ClearContents

chartheadrow = 3

For c = 2 To newtrendrow - 2 Step 2
    With ws_trends
        trendname = ws_trenddata.Range("A" & c).Value
        
        barcount = Len(trendname) - Len(Replace(trendname, "|", ""))
        
        If barcount = 0 Then x = 0 'And coderow = c
        If barcount = 1 Then x = 48 'And arearow = c
        If barcount = 2 Then x = 96 'And facilityrow = c
        
        Call FormatChartBlock(chartheadrow, trendname, x)
        y = Cells(chartheadrow + 1, 1).Top
        chartheadrow = chartheadrow + 14
        'ADD CHART
        .Shapes.AddChart(xlLineMarkers, x, y, 480, 180).Select
        With ActiveChart
            .SetSourceData Source:=ws_trenddata.Range(Cells(c, 2).Address, Cells(c + 1, trenddatacolumn).Address), PlotBy:=xlRows
            .SeriesCollection(1).XValues = ws_trenddata.Range(Cells(1, 2).Address, Cells(1, trenddatacolumn).Address)
            .HasLegend = True
            .Legend.Position = xlLegendPositionBottom
            .HasTitle = False
            .SeriesCollection(1).Name = "Model QTY"
            .SeriesCollection(1).MarkerStyle = 8
            .SeriesCollection(1).MarkerSize = 7
            .SeriesCollection(1).Format.Line.Weight = 2
            .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(68, 84, 106)
            .SeriesCollection(1).MarkerBackgroundColor = RGB(255, 255, 255)
            .SeriesCollection(2).Name = "Estimate QTY"
            .SeriesCollection(2).MarkerStyle = 8
            .SeriesCollection(2).MarkerSize = 7
            .SeriesCollection(2).Format.Line.Weight = 2
            .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(238, 150, 34)
            .SeriesCollection(2).MarkerBackgroundColor = RGB(255, 255, 255)
            .Axes(xlValue).HasTitle = True
            'Set cmt = Sheets("TrendData").Range("A" & c + 1).Comment.Text
            .Axes(xlValue).AxisTitle.Caption = "Quantity - " & ws_trenddata.Range("A" & c).Comment.Text
            .Axes(xlCategory).CategoryType = xlCategoryScale
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

'SORT TRENDDATA


With ws_compare
    For Each item In .Range(Cells(firstrow, 1).Address, Cells(lastrow, 1).Address)
        itemrow = item.Row
        'Set tcfound = Nothing
        If item.IndentLevel = 0 Then
            costcode = item
            Set tcfound = ws_trends.Cells.Find(What:=costcode, LookIn:=xlValues, SEARCHORDER:=xlByRows, MatchCase:=True)
            If Not tcfound Is Nothing Then
                setfound = tcfound.Address
                'CREATE HYPERLINKS BETWEEN EACH CHART AND ITEMROW WITHIN COMPARISON
                ws_compare.hyperlinks.Add Anchor:=ws_compare.Range("F" & itemrow), Address:="", SubAddress:="'Trends'!" & setfound
                ws_trends.hyperlinks.Add Anchor:=ws_trends.Range(setfound), Address:="", SubAddress:="'Comparison'!F" & itemrow
            End If
        End If
        If item.IndentLevel = 1 Then
            area = costcode & " | " & item
            Set tcfound = ws_trends.Cells.Find(What:=area, LookIn:=xlValues, SEARCHORDER:=xlByRows, MatchCase:=True)
            If Not tcfound Is Nothing Then
                setfound = tcfound.Address
                'CREATE HYPERLINKS BETWEEN EACH CHART AND ITEMROW WITHIN COMPARISON
                ws_compare.hyperlinks.Add Anchor:=ws_compare.Range("F" & itemrow), Address:="", SubAddress:="'Trends'!" & setfound
                ws_trends.hyperlinks.Add Anchor:=ws_trends.Range(setfound), Address:="", SubAddress:="'Comparison'!F" & itemrow
            End If
        End If
        If item.IndentLevel = 3 Then
            facility = costcode & " | " & area & " | " & item
            Set tcfound = ws_trends.Cells.Find(What:=facility, LookIn:=xlValues, SEARCHORDER:=xlByRows, MatchCase:=True)
            If Not tcfound Is Nothing Then
                setfound = tcfound.Address
                'CREATE HYPERLINKS BETWEEN EACH CHART AND ITEMROW WITHIN COMPARISON
                ws_compare.hyperlinks.Add Anchor:=ws_compare.Range("F" & itemrow), Address:="", SubAddress:="'Trends'!" & setfound
                ws_trends.hyperlinks.Add Anchor:=ws_trends.Range(setfound), Address:="", SubAddress:="'Comparison'!F" & itemrow
            End If
        End If
        Set tcfound = Nothing
    Next
    
End With



End Sub

Sub FormatChartBlock(chartheadrow, trendname, x)
If x = 0 Then
    With Sheets("Trends").Range("A" & chartheadrow & ":J" & chartheadrow)
        .MergeCells = True
        .Interior.color = 8421504
        .HorizontalAlignment = xlLeft
        '.Font.color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Value = trendname
    End With
    With Sheets("Trends").Range("A" & chartheadrow + 1 & ":J" & chartheadrow + 12)
        .MergeCells = True
        '.Border = True
    End With
End If
If x = 48 Then
    With Sheets("Trends").Range("B" & chartheadrow & ":K" & chartheadrow)
        .MergeCells = True
        .Interior.color = 10921638
        .HorizontalAlignment = xlLeft
        '.Font.color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Value = trendname
    End With
    With Sheets("Trends").Range("B" & chartheadrow + 1 & ":K" & chartheadrow + 12)
        .MergeCells = True
        '.Border = True
    End With
End If
If x = 96 Then
    With Sheets("Trends").Range("C" & chartheadrow & ":L" & chartheadrow)
        .MergeCells = True
        .Interior.color = 14277081
        .HorizontalAlignment = xlLeft
        '.Font.color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Value = trendname
    End With
    With Sheets("Trends").Range("C" & chartheadrow + 1 & ":L" & chartheadrow + 12)
        .MergeCells = True
        '.Border = True
    End With
End If
End Sub
