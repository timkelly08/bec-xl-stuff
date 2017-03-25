Attribute VB_Name = "TESTING"
Sub AddSht_AddCode()
    Dim wb As Workbook
    Dim xPro As VBIDE.VBProject
    Dim xCom As VBIDE.VBComponent
    Dim xMod As VBIDE.CodeModule
    Dim xLine As Long

    Set wb = ActiveWorkbook

    With wb
        Set xPro = .VBProject
        Set xCom = xPro.VBComponents("Sheet1")
        Set xMod = xCom.CodeModule

        With xMod
            xLine = .CreateEventProc("followhyperlink", "Worksheet")
            xLine = xLine + 1
            .InsertLines xLine, "   ActiveWindow.ScrollRow = ActiveCell.Row"
        End With
    End With

End Sub


Sub CONFIGsheet()
    With Sheets("CONFIG")
        If .Visible = True Then
            .Visible = False
        Else
            .Visible = True
            .Activate
        End If
    End With
End Sub
Sub configrules()
    With Sheets("RULES")
        If .Visible = True Then
            .Visible = False
        Else
            .Visible = True
            .Activate
            .Range("A1").Select
        End If
    End With
End Sub


Private Sub OKButton_Click()
Application.ScreenUpdating = False
begincheck = Timer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FIND ANY 2013 SPECIFIC FUNCTIONS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

main.GetAssHeaders

'set the project name on the setup page
Sheets("QTSetup").Range("N1").Value = "Project Name"
pmv = Sheets(AssembleData).Range("A1").Value
pmvsplit = Split(pmv, "-")
ProjectName = pmvsplit(0)
Sheets("QTSetup").Range("N2").Value = ProjectName

versionname = versionnametext.Text
ProjectName = projectnametext.Text
Sheets("QTSetup").Range("N1").Value = "Project Name"
Sheets("QTSetup").Range("N2").Value = ProjectName

Sheets("QTSetup").Range("M1").Value = "CombineChartsCheck"
If CombineChartsCheck.Value = True Then
    Sheets("QTSetup").Range("M2").Value = "True"
Else
    Sheets("QTSetup").Range("M2").Value = "False"
End If


Dim i
Dim x

For i = 0 To GroupListBox.ListCount - 1
    If GroupListBox.Selected(i) = True Then
        groupval = GroupListBox.List(i)
        If groupval = "" Then GoTo nextsheet
        
        sr = i + 2
        Sheets("QTSetup").Range("G" & sr).ClearComments
        Sheets("QTSetup").Range("G" & sr).AddComment
        Sheets("QTSetup").Range("G" & sr).Comment.Text Text:="1"
        
        main.TypeandInstanceCounts
        
        'ADD A NEW SHEET FROM THE ADDINS TEMPLATES
        ThisWorkbook.Sheets("DETAILS").Copy After:=ActiveWorkbook.Sheets(Sheets.count)
        
        'CHECK IF A SHEET NAME IS VALID, AND IF NOT RENAME IT
        SheetNameString = groupval
        Call main.IsValidSheetName(SheetNameString)
        newsht = SheetNameString
        
        With Sheets("Trends")
        .Name = newsht
        
        'Call Main.newtrendgroup
        .Range("A3").Value = ProjectName & ": Model Data Detail"
        .Range("A4").Value = groupval
        .Range("B6").Value = "TREND DATA"
        .Range("B7").Value = Sheets("QTSetup").Range("I1")

        .Range("C7").Value = versionname
        
        y = 176
        chartheadrow = 9
            row1 = 7
            For x = 0 To TrendListBox.ListCount - 1
            If TrendListBox.Selected(x) = True Then
            trendvalue = TrendListBox.List(x)
            If trendvalue = "" Then GoTo nexttrend
            
                    tr = x + 2
                    Sheets("QTSetup").Range("K" & tr).ClearComments
                    Sheets("QTSetup").Range("K" & tr).AddComment
                    Sheets("QTSetup").Range("K" & tr).Comment.Text Text:="1"
                
                
                checkval = Application.CountIfs(Sheets(AssembleData).Range(firstshtpropaddr & ":" & lastshtpropaddr), groupval, Sheets(AssembleData).Range(firsttrdpropaddr & ":" & lasttrdpropaddr), trendvalue)
                finalval = Application.SumIfs(Sheets(AssembleData).Range(firstquantityaddr & ":" & lastquantityaddr), Sheets(AssembleData).Range(firstshtpropaddr & ":" & lastshtpropaddr), groupval, Sheets(AssembleData).Range(firsttrdpropaddr & ":" & lasttrdpropaddr), trendvalue)
                
                If checkval <> 0 Then

                
                    row1 = row1 + 1
                    y = y + 15
                    chartheadrow = chartheadrow + 1
                    .Range("B" & row1).Value = TrendListBox.List(x)
                    .Range("C" & row1).Value = finalval
                End If
            End If
                        
nexttrend:
            Next x
            
            'Set data as a table
            thistable = "TrendData" & i
            .ListObjects.Add(xlSrcRange, Worksheets(newsht).Range("$B$7:$C$7", "$B$8:$C$" & row1), , xlYes).Name = thistable
            .ListObjects(1).TableStyle = "TableStyleMedium4"
            
            'reset column widths back to template
            .Range("B:B").ColumnWidth = 30
            .Range("C:J").ColumnWidth = 10
            
            'Add Charts to the sheet
            If BuildTrendReport.CombineChartsCheck.Value = True Then
                .Shapes.AddChart(xlLineMarkers, 14.25, y, 440, 180).Select
                ActiveChart.SetSourceData Source:=Range(thistable), PlotBy:=xlRows
                ActiveChart.SeriesCollection(1).XValues = "='" & newsht & "'!$C$7"
                            
                'Setup Chart block format
                trendname = "All Components"
                Call Format.FormatChartBlock(chartheadrow, trendname)
            Else
                Dim usetable As ListObject
                Set usetable = Sheets("Trends").ListObjects(thistable)
                For c = 1 To usetable.ListRows.count
                
                'GET THE UOM AND VERIFY THAT IT IS A SINGLE UOM
                prevUOM = ""
                For Each uom In Sheets(AssembleData).Range(firstunitaddr & ":" & lastunitaddr)
                    uomrow = uom.Row
                    groupcheck = Sheets(AssembleData).Cells(uomrow, sheetpropertycolumn).Value
                    If groupcheck = groupval Then
                        trendcheck = Sheets(AssembleData).Cells(uomrow, trendpropertycolumn).Value
                        If trendcheck = .Range("B" & 7 + c).Value Then
                            If prevUOM = "" Then
                                prevUOM = uom
                            Else
                                TrendUOM = uom
                                If TrendUOM <> prevUOM Then
                                    TrendUOM = "Invalid"
                                    GoTo continue2:
                                End If
                            End If
                        End If
                    End If
                Next uom
continue2:
                    
                    'ADD THE CHART
                    .Shapes.AddChart(xlLineMarkers, 14.25, y, 440, 180).Select
                    With ActiveChart
                        .SetSourceData Source:=usetable.ListRows(c).Range, PlotBy:=xlRows
                        .SeriesCollection(1).XValues = "='" & newsht & "'!$C$7"
                        .HasLegend = False
                        .SeriesCollection(1).MarkerStyle = 8
                        .SeriesCollection(1).MarkerSize = 7
                        .SeriesCollection(1).Format.Line.Weight = 2
                        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(68, 84, 106)
                        .SeriesCollection(1).MarkerBackgroundColor = RGB(255, 255, 255)
                        .Axes(xlValue).HasTitle = True
                        .Axes(xlValue).AxisTitle.Caption = "Quantity (" & TrendUOM & ")"
                        .HasTitle = False
                        '.Axes(xlValue).AxisTitle.Font.Name = "bookman"
                        '.Axes(xlValue).AxisTitle.Font.Size = 10
                        '.Axes(xlValue).AxisTitle.Characters(9, 5).Font.Italic = True
                    End With
                    'Setup Chart block format
                    trendname = .Range("B" & c + 7).Value
                    Call Format.FormatChartBlock(chartheadrow, trendname)
                    y = y + 210
                    chartheadrow = chartheadrow + 14
                Next
            End If
        End With
    End If
nextsheet:
Next i

MacroName = "Build Quantity Trend"
groupcount = Sheets.count - 3
main.SendAudit
    
Application.DisplayAlerts = False
Sheets(AssembleData).Delete
Application.DisplayAlerts = True

'Worksheets("SUMMARY").Activate
Application.ScreenUpdating = True
BuildTrendReport.Hide
finishcheck = Timer
totaltime = finishcheck - begincheck
Debug.Print totaltime
End Sub



Sub GetAssHeaders()

AssembleData = Sheets("Trends").Name
AssembleHeaderRow = Application.Match("Color", Sheets("Trends").Range("A:A"), False)
sourceidcolumn = Application.Match("Source ID", Worksheets(AssembleData).Range("A" & AssembleHeaderRow & ":AAA" & AssembleHeaderRow), False)
quantitycolumn = Application.Match("Quantity", Worksheets(AssembleData).Range("A" & AssembleHeaderRow & ":AAA" & AssembleHeaderRow), False)
unitcolumn = Application.Match("Unit", Worksheets(AssembleData).Range("A" & AssembleHeaderRow & ":AAA" & AssembleHeaderRow), False)
itemcolumn = Application.Match("Item", Worksheets(AssembleData).Range("A" & AssembleHeaderRow & ":AAA" & AssembleHeaderRow), False)
sheetpropertycolumn = Sheets("QTSetup").Range("E2").Value + 1
trendpropertycolumn = Sheets("QTSetup").Range("I2").Value + 1
lastrow = Worksheets(AssembleData).Cells(Rows.count, sourceidcolumn).End(xlUp).Row
firstidaddr = Worksheets(AssembleData).Cells(AssembleHeaderRow + 1, sourceidcolumn).Address
lastidaddr = Worksheets(AssembleData).Cells(lastrow, sourceidcolumn).Address
firstquantityaddr = Worksheets(AssembleData).Cells(AssembleHeaderRow + 1, quantitycolumn).Address
lastquantityaddr = Worksheets(AssembleData).Cells(lastrow, quantitycolumn).Address
firstshtpropaddr = Worksheets(AssembleData).Cells(AssembleHeaderRow + 1, sheetpropertycolumn).Address
lastshtpropaddr = Worksheets(AssembleData).Cells(lastrow, sheetpropertycolumn).Address
firsttrdpropaddr = Worksheets(AssembleData).Cells(AssembleHeaderRow + 1, trendpropertycolumn).Address
lasttrdpropaddr = Worksheets(AssembleData).Cells(lastrow, trendpropertycolumn).Address
firstitemaddr = Worksheets(AssembleData).Cells(AssembleHeaderRow + 1, itemcolumn).Address
lastitemaddr = Worksheets(AssembleData).Cells(lastrow, itemcolumn).Address
firstunitaddr = Worksheets(AssembleData).Cells(AssembleHeaderRow + 1, unitcolumn).Address
lastunitaddr = Worksheets(AssembleData).Cells(lastrow, unitcolumn).Address
LastAssRow = lastrow

End Sub
Sub getcolor()
Debug.Print ActiveCell.Font.color
End Sub

Sub TypeandInstanceCounts()
startcount = Timer
    'GET A COUNT OF INSTANCES IN GROUP
    InstanceCount = Application.CountIf(Sheets(AssembleData).Range(firstshtpropaddr & ":" & lastshtpropaddr), groupval)
    'GET A COUNT OF TYPES IN GROUP
    Dim types As New Collection
    For Each itm In Sheets(AssembleData).Range(firstshtpropaddr & ":" & lastshtpropaddr)
        If itm = groupval Then
            itmrow = itm.Row
            itemname = Sheets(AssembleData).Cells(itmrow, itemcolumn).Value
            
            On Error Resume Next
            types.Add itemname, Chr(34) & itemname & Chr(34)
        End If
    Next itm
    On Error GoTo 0
    TypeCount = types.count
    
    Do Until types.count = 0
        types.Remove (1)
    Loop
endcount = Timer

Debug.Print "Type and Instances - "; endcount - startcount

End Sub


Sub FormatTrendSheet()
With Sheets("Trends")
    .Range("A:A").ColumnWidth = 3
    .Range("A:A").RowHeight = 15
    '.Range("B:B").ColumnWidth = 15
    .Range("B:K").ColumnWidth = 11
    .Range("L:L").ColumnWidth = 3
    With .Range("A1:L1")
        .MergeCells = True
        .Interior.color = 14277081
        .RowHeight = 40
    End With
    With .Range("A2:L2")
        .MergeCells = True
        .Interior.color = 6968388
        .RowHeight = 4
    End With
    With .Range("A3:L3")
        .MergeCells = True
        .Font.color = -9808828
        .RowHeight = 32
        .HorizontalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 18
    End With
    With .Range("A4:L4")
        .MergeCells = True
        .Font.color = -9808828
        .Font.Bold = True
        .RowHeight = 25
        .HorizontalAlignment = xlCenter
        .Font.Name = "Calibri"
        .Font.Size = 18
    End With
    With .Range("A5:L5")
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    With .Range("B6:K6")
        .MergeCells = True
        .Interior.color = 6968388
        .HorizontalAlignment = xlLeft
        .Font.color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 12
    End With
    
End With

End Sub

Sub FormatChartBlock(chartheadrow, trendname)
With Sheets("Trends").Range("B" & chartheadrow & ":G" & chartheadrow)
    .MergeCells = True
    .Interior.color = 6968388
    .HorizontalAlignment = xlLeft
    .Font.color = RGB(255, 255, 255)
    .Font.Bold = True
    .Font.Name = "Calibri"
    .Font.Size = 12
    .Value = trendname
End With
With Sheets("Trends").Range("H" & chartheadrow & ":J" & chartheadrow)
    .MergeCells = True
    .Interior.color = 6968388
    .HorizontalAlignment = xlLeft
    .Font.color = RGB(255, 255, 255)
    .Font.Bold = True
    .Font.Name = "Calibri"
    .Font.Size = 12
    .Value = "Comments"
End With
With Sheets("Trends").Range("B" & chartheadrow + 1 & ":G" & chartheadrow + 12)
    .MergeCells = True
    '.Border = True
End With
With Sheets("Trends").Range("H" & chartheadrow + 1 & ":J" & chartheadrow + 12)
    .Merge (True)
    .HorizontalAlignment = xlLeft
    .Font.Name = "Calibri"
    .Font.Size = 11
    '.Border = True
End With
End Sub



Sub SummaryCOSTCODEComparisontest()
    
'DECLARE AND SET WORKBOOKS AND WORKSHEETS
    Dim wb_master As Workbook
    Dim ws_config As Worksheet
    Dim ws_creport As Worksheet
    Dim ws_compare As Worksheet
    Dim ws_QTOflat As Worksheet
    Dim ws_trenddata As Worksheet
    Dim ws_trends As Worksheet
    Dim ws_assembly As Worksheet

    Set wb_master = ActiveWorkbook
    Set ws_config = Worksheets("CONFIG")
    Set ws_creport = Worksheets("Cost Report")
    Set ws_compare = Worksheets("Comparison")
    Set ws_QTOflat = Worksheets("MasterQTO_flat")
    Set ws_trenddata = Worksheets("TrendData")
    Set ws_trends = Worksheets("Trends")
    Set ws_assembly = Worksheets("Assembly Codes & Unit Costs")

'DECLARE AND SET TABLES
    Dim lo_qto As ListObject
    Dim lo_compare As ListObject
    Dim lo_trend As ListObject
    Dim lo_tolerance As ListObject
    
    Set lo_qto = ws_config.ListObjects("QTO_CONFIG")
    Set lo_compare = ws_config.ListObjects("COMPARE_CONFIG")
    Set lo_tolerance = ws_config.ListObjects("TOLERANCE_CONFIG")


'DECLARE AND SET COLUMNS


'DECLARE AND SET ROWS


'DECLARE AND SET OTHERS'

    Dim areacol As Integer
    Dim allareas As Integer
    Dim facilitycol As Integer
    Dim assemblycol As Integer
    Dim itemcol As Integer
    Dim sheet As Variant
    Dim area As Variant
    Dim facility As Variant
    Dim assemblycode As Variant
    Dim item As Variant
    Dim lastcol As Integer
    Dim lastrow As Long
    Dim QTOflat_lastrow As Long
    Dim lastitem As String
    Dim itemrow As Variant
    Dim pArea As String
    Dim pFacility As String
    Dim pAssemblyCode As String
    Dim pItem As String
    Dim columnheaders() As String
    Dim QTO_flatExists As Boolean
    Dim CompareConfig As Variant
    Dim isdimensioned As Boolean
    Dim x As Integer
    Dim y As Integer
    Dim lastheader As String
    Dim pAreaRow As Variant
    Dim pFacilityRow As Variant
    Dim pAssemblyRow As Variant
    Dim pItemRow As Variant
    Dim itemvalue As String
    Dim itemvaluecol As Integer
    Dim costcodecol As Integer
    Dim costcode As Variant
    Dim assemblycodeformula As String
    Dim itemformula As String
    Dim MATLcode As String
    Dim MLcode As String
    Dim SCcode As String
    Dim coderow As Variant
    Dim searchnext As Boolean
    Dim currentUOM As String
    Dim acrossheaders() As Variant
    Dim uom As String
    Dim uomrow As Integer
    Dim code As Variant
    Dim deviationcolor As ColorScale
    Dim codedesc As String
    Dim reportcode As Variant
    Dim foundinreport As Variant
    Dim combinedvalfound As Boolean
    Dim alreadyincomparison As Variant
    Dim reportArea As Variant
    Dim reportFacility As Variant
    Dim reportFD As Variant
    Dim reportCC As Variant
    Dim reportCCD As Variant
    Dim reportResource As Variant
    Dim reportUOM As Variant
    Dim reportQTY As Variant
    Dim lastreportcode As Variant
    Dim reportcombined As Variant
    Dim lAlpha As Variant
    Dim lRemainder As Variant
    Dim reportLastcol As Integer
    Dim reportnotescol As Integer
    Dim QTOflat_qcol As Integer
    Dim qtycol As Integer
    Dim codecol As Integer
    Dim lastcomparerow As Long
    Dim compareareacol As Integer


    'VERIFY THAT THE MASTERQTO_FLAT SHEET EXISTS
    For Each sheet In wb_master.Sheets
        If sheet.Name = ws_QTOflat.Name Then
            QTO_flatExists = True
        End If
    Next
    If QTO_flatExists = False Then
        MsgBox "It looks like you have not imported QEX or MTO files. Please import data by using the 'Combine QEX Files' command and selecting QEX or MTO files."
        Application.ScreenUpdating = True
        Exit Sub
    End If

End Sub

Sub makelinks()
    With Sheets("Comparison")
        .hyperlinks.Add Anchor:=.Range("F4"), Address:="", SubAddress:="'Trends'!A113"
        .hyperlinks.Add Anchor:=Sheets("Trends").Range("A113"), Address:="", SubAddress:="'Comparison'!F4"
    End With
End Sub

Sub AddUOMcomment()
With Sheets("TrendData (2)")
    For x = 2 To 139
        .Cells(x, 1).ClearComments
        .Cells(x, 1).AddComment
        .Cells(x, 1).Comment.Text Text:=.Cells(x, 2).Value
    Next
End With
End Sub


Sub test()

Dim rule As MSXML2.DOMDocument
strXML = ActiveCell.Value

    Set rule = New MSXML2.DOMDocument

    If Not rule.LoadXML(strXML) Then  'strXML is the string with XML'
        Err.Raise rule.parseError.ErrorCode, , rule.parseError.reason
    End If

'Dim point As IXMLDOMNode
'Set point = rule.FirstChild

'Debug.Print point.SelectSingleNode("X").Text
'Debug.Print point.SelectSingleNode("Y").Text

'For Each Child In point.ChildNodes
    'Debug.Print point.ChildNodes.Length
    RuleName = rule.FirstChild.ChildNodes(0).Text
    uom = rule.FirstChild.ChildNodes(2).Text
    costcode = rule.FirstChild.ChildNodes(3).Text
    Formula1 = rule.FirstChild.ChildNodes(4).Text
    replace1 = rule.FirstChild.ChildNodes(5).Text
    
      
    Debug.Print replace1
'Next
End Sub




'<Rules><CostCode><Value>12</Value><Value>14</Value></CostCode><Assembly>This One</Assembly></Rules>
