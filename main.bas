Attribute VB_Name = "main"
Option Explicit
Public overwriteQTO As Boolean

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OPEN FILE DIALOG TO GET MULTIPLE QEX FILES LOADED INTO QTO REPORT
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub openQEXfiles()

    Application.ScreenUpdating = False

    Dim QEXlist As Long
    Dim QEXname As String
    Dim openQEX As Workbook
    Dim sheettocheck As String
    Dim sheetvalid As Variant
    Dim sheet As Worksheet
    Dim masterQTOFile As Workbook
    Dim startposition As Long
    Dim firstrow As Long
    Dim lastrow As Long
    Dim itemcolumn As Long
    Dim MasterQTO_flatExists As Boolean
    Dim lastcolumnheader As Long
    Dim columnheaders As Variant
    Dim requiredcolumn As Variant
    Dim x As Integer
    Dim columnmatch As Integer
    Dim isdimensioned As Boolean
    Dim lastcolumn As Integer
    Dim pastecolumn As Integer
    Dim pasteaddress As String
    Dim takeofftypecol As Integer
    Dim takeofftypestart As String
    Dim takeofftypeend As String
    Dim lvl3codecol As Integer
    Dim lvl3codestart As String
    Dim lvl3codeend As String
    Dim costcodecol As Integer
    Dim firstcode As Variant
    Dim firstposition As Long
    Dim startborder As String
    Dim endborder As String
    Dim requiredcolumnshort As Variant
    Dim requiredcolumn2short As Variant
    Dim thiscolumnheader As Variant
    Dim thiscolumnheadershort As Variant
    Dim totalrows As Variant
    Dim IsCivil3D As Variant
    Dim requiredcolumn2 As Variant
    Dim isreq As Variant
    
    
    
    Set masterQTOFile = ActiveWorkbook
    
    For Each sheet In masterQTOFile.Sheets
    If sheet.Name = "MasterQTO_flat" Then
        MasterQTO_flatExists = True
    End If
Next
    
    
    'OPEN THE FILE DIALOG
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        If .Show <> -1 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        'IF THE MASTERQTO_FLAT SHEET DOES EXIST, ASK USER IF THEY WANT TO OVERWRITE OR APPEND. IF OVERWRITE, DELETE SHEET AND UPDATE FLAT_EXISTS
        If MasterQTO_flatExists = True Then QEXimport.Show
        If overwriteQTO = True Then
            Application.DisplayAlerts = False
            Sheets("MasterQTO_flat").Delete
            Application.DisplayAlerts = True
            MasterQTO_flatExists = False
        End If
            
        
         
        'IF THE MASTERQTO_FLAT SHEET DOES NOT EXIST, CREATE ONE
        If MasterQTO_flatExists <> True Then
            Worksheets.Add().Name = "MasterQTO_flat"
            lastcolumnheader = Sheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Rows.count
            columnheaders = Sheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Columns(2)
            
            With Sheets("MasterQTO_flat")
                With .Range("A1")
                    .Value = "EVLLRT-All Areas QTO Report"
                    .Font.Bold = True
                    .Font.Size = 14
                End With
                With .Range(Cells(3, 1), Cells(3, lastcolumnheader))
                    .Value = Application.Transpose(columnheaders)
                    .Font.color = 16777215
                    .Font.Bold = True
                    .Interior.color = 12419407
                    With .Borders
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = xlAutomatic
                    End With
                End With
            End With
        End If
            
        'SET THE POSITION TO START IMPORTING QEX AND MTO FILES INTO THE MASTERQTO_FLAT SHEET
        startposition = Sheets("MasterQTO_flat").Cells(Rows.count, 7).End(xlUp).Row + 1
        firstposition = startposition
         
         
        'OPEN EACH FILE
        For QEXlist = 1 To .SelectedItems.count
            QEXname = .SelectedItems(QEXlist)
            Set openQEX = Workbooks.Open(Filename:=QEXname)
                
                'VERIFY THE SOURCE OF THE EXCEL FILE
                For Each sheet In openQEX.Sheets
                    If sheet.Visible = True Then
                        sheettocheck = sheet.Name
                        sheetvalid = Application.Match("Model Version ID", Sheets(sheettocheck).Range("A:A"), False)
                        If IsError(sheetvalid) Then
                            sheetvalid = Application.Match("Color", Sheets(sheettocheck).Range("A:A"), False)
                            If IsError(sheetvalid) Then
                                If Not Sheets(sheettocheck).Range("A1").Value = "EVLLRT - Manual Takeoffs" Then
                                    MsgBox openQEX.Name & " does not look like a QEX or MTO file and will not be imported.", vbOKOnly, "Import Error"
                                    Application.DisplayAlerts = False
                                    openQEX.Close
                                    Application.DisplayAlerts = True
                                    GoTo NEXTBOOK
                                End If
                            End If
                        End If

                        'CHECK TO SEE IF THIS IS AN MTO FILE OR QEX AND COPY THE VALUES
                        If Sheets(sheettocheck).Range("A1").Value = "EVLLRT - Manual Takeoffs" Then
                            
                            'COPY DATA FROM OPEN MTO TO MasterQTO_flat SHEET
                            With Worksheets(sheettocheck)
                                lastcolumn = .Cells(3, Columns.count).End(xlToLeft).Column
                                itemcolumn = Application.Match("Item", .Range(Cells(3, 1), Cells(3, lastcolumn)), False)
                                firstrow = 4
                                lastrow = .Cells(Rows.count, itemcolumn).End(xlUp).Row
                                For Each columnheaders In .Range(Cells(3, 1), Cells(3, lastcolumn))
                                    pastecolumn = Application.Match(columnheaders, masterQTOFile.Sheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns("Column Name").DataBodyRange, False)
                                    pasteaddress = Cells(startposition, pastecolumn).Address
                                    .Range(Cells(firstrow, columnheaders.Column), Cells(lastrow, columnheaders.Column)).Copy masterQTOFile.Sheets("MasterQTO_flat").Range(pasteaddress)
                                Next
                                takeofftypecol = Application.Match("TE_QTO_Method", masterQTOFile.Sheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns("Column Name").DataBodyRange, False)
                                lvl3codecol = Application.Match("Assembly Code Level 3", masterQTOFile.Sheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns("Column Name").DataBodyRange, False)
                                costcodecol = Application.Match("Assembly Code", masterQTOFile.Sheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns("Column Name").DataBodyRange, False)
                                takeofftypestart = .Cells(startposition, takeofftypecol).Address
                                takeofftypeend = .Cells(startposition + (lastrow - firstrow), takeofftypecol).Address
                                lvl3codestart = .Cells(startposition, lvl3codecol).Address
                                lvl3codeend = .Cells(startposition + (lastrow - firstrow), lvl3codecol).Address
                                firstcode = .Cells(startposition, costcodecol).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                                With masterQTOFile.Sheets("MasterQTO_flat")
                                .Range(takeofftypestart & ":" & takeofftypeend).Value = "Manual"
                                .Range(lvl3codestart & ":" & lvl3codeend).Value = "=IFNA(VLOOKUP(" & firstcode & ", 'Cost Report'!G:I,3,FALSE),""Not Assigned"")"
                                    startborder = Cells(startposition, 1).Address
                                    lastcolumnheader = masterQTOFile.Sheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Rows.count
                                    endborder = Cells(startposition + (lastrow - firstrow), lastcolumnheader).Address
                                    With .Range(startborder & ":" & endborder)
                                        With .Borders
                                            .LineStyle = xlContinuous
                                            .Weight = xlThin
                                            .ColorIndex = xlAutomatic
                                        End With
                                    End With
                                End With
                            End With
    
                            startposition = startposition + (lastrow - firstrow) + 1
                    
                            Application.DisplayAlerts = False
                            openQEX.Close
                            Application.DisplayAlerts = True
                        Else
                            'VERIFY THE QEX FILES HAVE THE CORRECT COLUMN HEADERS
                            For Each isreq In masterQTOFile.Sheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns("IsRequired?").DataBodyRange
                                If isreq = "True" Then
                                requiredcolumn = masterQTOFile.Sheets("CONFIG").Cells(isreq.Row, 2).Value
                                columnmatch = isreq.Row - 2
                                    'REMOVE UOM FROM QUANTITY COLUMNS SINCE THERE IS INCONSISTENT USE OF THE ()
                                    If InStr(requiredcolumn, "(") > 0 Then
                                        requiredcolumnshort = Split(requiredcolumn, " (")
                                        requiredcolumn = requiredcolumnshort(0)
                                    End If
                                    thiscolumnheader = Sheets(sheettocheck).Cells(sheetvalid, columnmatch).Value
                                    If InStr(thiscolumnheader, "(") > 0 Then
                                        thiscolumnheadershort = Split(thiscolumnheader, " (")
                                        thiscolumnheader = thiscolumnheadershort(0)
                                    End If
                                    If thiscolumnheader <> requiredcolumn Then
                                        IsCivil3D = MsgBox("During the import, we found some columns that don't match within:" & vbCrLf & vbCrLf & openQEX.Name & vbCrLf & vbCrLf & "Is this an AutoCAD Civil 3D model?", vbYesNo, "Import Check")
                                        If IsCivil3D = vbYes Then
                                            'CHECK EACH COLUMN AND MATCH TO APPROPRIATE COLUMN IN MASTERQTO
                                            With Worksheets(sheettocheck)
                                                For Each requiredcolumn2 In masterQTOFile.Sheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns("Column Name").DataBodyRange
                                                    columnmatch = requiredcolumn2.Row - 2
                                                    thiscolumnheader = Application.Match(requiredcolumn2, .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                                    If IsError(thiscolumnheader) Then
                                                        'REMOVE UOM FROM QUANTITY COLUMNS SINCE THERE IS INCONSISTENT USE OF THE ()
                                                        If InStr(requiredcolumn2, "(") > 0 Then
                                                            requiredcolumn2short = Split(requiredcolumn2, " (")
                                                            requiredcolumn2 = requiredcolumn2short(0)
                                                        End If
                                                    End If
                                                    thiscolumnheader = Application.Match(requiredcolumn2, .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                                    If requiredcolumn2 = "TE_Area" Then
                                                        thiscolumnheader = Application.Match("Zone/Area (Assemble Property)", .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                                    End If
                                                    If requiredcolumn2 = "TE_Facility" Then
                                                        thiscolumnheader = Application.Match("Location (Assemble Property)", .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                                    End If
                                                    If requiredcolumn2 = "Assembly Code Level 1" Then
                                                        thiscolumnheader = Application.Match("Assembly Code", .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                                    End If
                                                    If requiredcolumn2 = "Assembly Code Level 2" Then
                                                        thiscolumnheader = Application.Match("Assembly Code", .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                                    End If
                                                    If requiredcolumn2 = "Assembly Code Level 3" Then
                                                        thiscolumnheader = Application.Match("Assembly Code", .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                                    End If
                                                    If requiredcolumn2 = "TE_Milestone" Then
                                                        thiscolumnheader = Application.Match("Activity ID (Assemble Property)", .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                                    End If
                                                    If requiredcolumn2 = "TE_Discipline" Then
                                                        thiscolumnheader = Application.Match("VE Option (Assemble Property)", .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                                    End If
                                                    If requiredcolumn2 = "TE_QTO_Method" Then
                                                        itemcolumn = Application.Match("Item", .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                                        firstrow = sheetvalid + 1
                                                        lastrow = .Cells(Rows.count, itemcolumn).End(xlUp).Row
                                                        totalrows = lastrow - firstrow
                                                        masterQTOFile.Sheets("MasterQTO_flat").Range(Cells(startposition, columnmatch).Address, Cells(startposition + totalrows, columnmatch).Address).Value = "Model"
                                                    End If
                                                    If requiredcolumn2 = "TE_QTO_Source" Then
                                                        itemcolumn = Application.Match("Item", .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                                        firstrow = sheetvalid + 1
                                                        lastrow = .Cells(Rows.count, itemcolumn).End(xlUp).Row
                                                        totalrows = lastrow - firstrow
                                                        masterQTOFile.Sheets("MasterQTO_flat").Range(Cells(startposition, columnmatch).Address, Cells(startposition + totalrows, columnmatch).Address).Value = "Civil 3D"
                                                    End If
                                                    If Not IsError(thiscolumnheader) Then
                                                        itemcolumn = Application.Match("Item", .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                                        firstrow = sheetvalid + 1
                                                        lastrow = .Cells(Rows.count, itemcolumn).End(xlUp).Row
                                                        .Range(Cells(firstrow, thiscolumnheader).Address, Cells(lastrow, thiscolumnheader).Address).Copy masterQTOFile.Sheets("MasterQTO_flat").Cells(startposition, columnmatch)
                                                    Else
                                                        'DRAW BORDERS
                                                        itemcolumn = Application.Match("Item", .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                                        firstrow = sheetvalid + 1
                                                        lastrow = .Cells(Rows.count, itemcolumn).End(xlUp).Row
                                                        With masterQTOFile.Sheets("MasterQTO_flat").Range(Cells(startposition, columnmatch).Address, Cells(startposition + lastrow - firstrow, columnmatch).Address).Borders
                                                            .LineStyle = xlContinuous
                                                            .Weight = xlThin
                                                            .ColorIndex = xlAutomatic
                                                        End With
                                                    End If
                                                Next
                                            End With
                                            startposition = startposition + (lastrow - firstrow) + 1
                                            Application.DisplayAlerts = False
                                            openQEX.Close
                                            Application.DisplayAlerts = True
                                            GoTo NEXTBOOK
                                        Else
                                            MsgBox "OK, then were missing some columns in the QEX file, starting with " & requiredcolumn & "." & vbCrLf & vbCrLf & _
                                            openQEX.Name & "will not be imported." & vbCrLf & vbCrLf & _
                                            "Please review the required columns within the 'Import Quantities Settings' table and import once all of the required columns are included within your QEX file."
                                            Application.DisplayAlerts = False
                                            openQEX.Close
                                            Application.DisplayAlerts = True
                                            GoTo NEXTBOOK
                                        End If
                                    End If
                                End If
                            Next
                        
                            'COPY DATA FROM OPEN QEX TO MasterQTO_flat SHEET
                            With Worksheets(sheettocheck)
                            
                            
                            
                            
                            
                            
                            'MADE ADJUSTMENT TO COPY ONE COLUMN AT A TIME FOR MEMORY ISSUE:
                            For Each requiredcolumn2 In masterQTOFile.Sheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns("Column Name").DataBodyRange
                                columnmatch = requiredcolumn2.Row - 2
                                thiscolumnheader = Application.Match(requiredcolumn2, .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                If IsError(thiscolumnheader) Then
                                    'REMOVE UOM FROM QUANTITY COLUMNS SINCE THERE IS INCONSISTENT USE OF THE ()
                                    If InStr(requiredcolumn2, "(") > 0 Then
                                        requiredcolumn2short = Split(requiredcolumn2, " (")
                                        requiredcolumn2 = requiredcolumn2short(0)
                                    End If
                                End If
                                thiscolumnheader = Application.Match(requiredcolumn2, .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                
                                If Not IsError(thiscolumnheader) Then
                                    itemcolumn = Application.Match("Item", .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                    firstrow = sheetvalid + 1
                                    lastrow = .Cells(Rows.count, itemcolumn).End(xlUp).Row
                                    .Range(Cells(firstrow, thiscolumnheader).Address, Cells(lastrow, thiscolumnheader).Address).Copy masterQTOFile.Sheets("MasterQTO_flat").Cells(startposition, columnmatch)
                                Else
                                    'DRAW BORDERS
                                    itemcolumn = Application.Match("Item", .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
                                    firstrow = sheetvalid + 1
                                    lastrow = .Cells(Rows.count, itemcolumn).End(xlUp).Row
                                    With masterQTOFile.Sheets("MasterQTO_flat").Range(Cells(startposition, columnmatch).Address, Cells(startposition + lastrow - firstrow, columnmatch).Address).Borders
                                        .LineStyle = xlContinuous
                                        .Weight = xlThin
                                        .ColorIndex = xlAutomatic
                                    End With
                                End If
                            Next
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            
'                                itemcolumn = Application.Match("Item", .Range("A" & sheetvalid & ":AAA" & sheetvalid), False)
'                                firstrow = sheetvalid + 1
'                                lastrow = .Cells(Rows.count, itemcolumn).End(xlUp).Row
'                                .Rows(firstrow & ":" & lastrow).Copy masterQTOFile.Sheets("MasterQTO_flat").Range("A" & startposition)
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            End With
    
                            startposition = startposition + (lastrow - firstrow) + 1
                    
                            Application.DisplayAlerts = False
                            openQEX.Close
                            Application.DisplayAlerts = True
                        End If
                    End If
                Next sheet
NEXTBOOK:
        Next QEXlist
 
    End With
    
    'VERIFY THAT SOMETHING WAS ACTUALLY IMPORTED, AND IF SO, REMOVE ANY GROUPING ROWS FROM THE IMPORT
    If firstposition <> startposition Then
        Call removegroupbys(firstposition)
    End If
    
    'ADD CONDITIONAL FORMATTING TO GROUP COLUMNS
    Sheets("MasterQTO_flat").Cells.FormatConditions.Delete
    With Sheets("MasterQTO_flat").Range("B:F").FormatConditions _
        .Add(xlCellValue, xlEqual, "Not Assigned")
        .Interior.color = 192
    End With
    
    With Sheets("MasterQTO_flat").UsedRange
        .WrapText = False
        .EntireColumn.AutoFit
        .EntireRow.AutoFit
        .ClearOutline
        .Font.color = 0
'        With Range("A4:" & Cells(lastcolumn, startposition - 1).Address).Borders
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
    End With
    If Not lastcolumnheader = 0 Then
        With Sheets("MasterQTO_flat").Range(Cells(3, 1), Cells(3, lastcolumnheader))
            .Font.color = 16777215
        End With
    End If
    
    Application.ScreenUpdating = True

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'RUN RULES AGAINST MASTER QTO REPORT AND ADD LINE ITEMS FOR EACH RULE MATCH
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub runrules()

Application.ScreenUpdating = False


    Dim lastrule As Long
    Dim rule As Variant
    Dim newitemrow As Long
    Dim rulerow As Long
    Dim newformulaname As String
    Dim newcostcode As String
    Dim newpropertyname As String
    Dim newpropertyvalue As String
    Dim newuom As String
    Dim newformula As String
    Dim propertycolumn As Long
    Dim firstvaladdr As String
    Dim lastvaladdr As String
    Dim loc As Range
    Dim firstfound As Long
    Dim foundnext As Long
    Dim formulabreakdown As Variant
    Dim i As Integer
    Dim draftformula As String
    Dim isQuantity As Variant
    Dim finalformula As String
    Dim quantitycol As String
    Dim assemblycodecol As Integer
    Dim milestonecol As Integer
    Dim methodcol As Integer
    Dim lastareacol As Integer
    Dim lastfaccol As Integer
    Dim replaceQTY As Boolean
    Dim starttime As Variant
    Dim finishtime As Variant
    
    'starttime = Time
    
    Sheets("MasterQTO_flat").Activate
    
    lastrule = Worksheets("Rules").Cells(Rows.count, 1).End(xlUp).Row
    newitemrow = Worksheets("MasterQTO_flat").Cells(Rows.count, 7).End(xlUp).Row + 1
    foundnext = 0
    assemblycodecol = Application.Match("Assembly Code", Sheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns("Column Name").DataBodyRange, False)
    milestonecol = Application.Match("TE_Milestone", Sheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns("Column Name").DataBodyRange, False)
    methodcol = Application.Match("TE_QTO_Method", Sheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns("Column Name").DataBodyRange, False)
    'Reverse search needed
    lastareacol = Application.Match("TE_Area", Sheets("CONFIG").Range("B9:B50"), False) + 6
    lastfaccol = Application.Match("TE_Facility", Sheets("CONFIG").Range("B9:B50"), False) + 6
    
    
    
    'STEP THROUGH EACH RULE AND COMPARE PROPERTY VALUES AGAINST MASTER QTO
    For Each rule In Sheets("Rules").Range("A2:A" & lastrule)
        With Sheets("Rules")
            rulerow = rule.Row
            newformulaname = .Cells(rulerow, 1).Value
            newcostcode = .Cells(rulerow, 2).Value
            newpropertyname = .Cells(rulerow, 3).Value
            newpropertyvalue = .Cells(rulerow, 4).Value
            newuom = .Cells(rulerow, 5).Value
            newformula = .Cells(rulerow, 6).Value
            replaceQTY = .Cells(rulerow, 7).Value
            newformula = Replace(newformula, "*", "|")
        End With
        
        'BREAK DOWN THE FORMULA AND COMPOSE A DRAFT TO POINT TO THE CORRECT CELLS AND VALUES
        formulabreakdown = Replace(newformula, "[", "~")
        formulabreakdown = Replace(formulabreakdown, "]", "~")
        formulabreakdown = Split(formulabreakdown, "~")
        draftformula = "=iferror("
        For i = 1 To UBound(formulabreakdown)
            isQuantity = Application.Match(formulabreakdown(i), Sheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Columns(2), False)
            If Not IsError(isQuantity) Then
                quantitycol = Split(Cells(1, isQuantity).Address, "$")(1)
                draftformula = draftformula & quantitycol & "[rowholder]"
            Else
                draftformula = draftformula & formulabreakdown(i)
            End If
        Next
        draftformula = Replace(draftformula, "|", "*")
        draftformula = draftformula & ",0)"
        
        'WITHIN MASTER QTO, FIND ANY MATCHES WITHIN THE PROPERTY COLUMN STEP TO THE NEXT
        With Sheets("MasterQTO_flat")
            propertycolumn = Application.Match(newpropertyname, .Range("A3:CC3"), False)
            firstvaladdr = .Cells(4, propertycolumn).Address
            lastvaladdr = .Cells(newitemrow - 1, propertycolumn).Address
            With Range(firstvaladdr & ":" & lastvaladdr)
                Set loc = .Cells.Find(What:=newpropertyvalue)
                If Not loc Is Nothing Then
                    firstfound = loc.Row
                    Do Until foundnext = firstfound
                        If replaceQTY = True Then
                            finalformula = Replace(draftformula, "[rowholder]", loc.Row)
                            'Range("B" & newitemrow).Value = Range("B" & loc.Row).Value
                            'Range("C" & newitemrow).Value = Range("C" & loc.Row).Value
                            .Range("F" & loc.Row).Value = "=IFNA(VLOOKUP(""" & newcostcode & """, 'Cost Report'!G:I,3,FALSE),""NOT ASSIGNED"")"
                            'Range("G" & newitemrow).Value = newformulaname
                            'Range("H" & newitemrow).Value = "Extrapolated from " & Range("H" & loc.Row).Value
                            'Range("I" & newitemrow).Value = Range("I" & loc.Row).Value & "-E"
                            .Range("J" & loc.Row).Value = finalformula
                            .Range("K" & loc.Row).Value = newuom
                            With Cells(loc.Row, assemblycodecol)
                                .NumberFormat = "@"
                                .Value = newcostcode
                            End With
                            'Cells(newitemrow, milestonecol).Value = Cells(loc.Row, milestonecol).Value
                            'Cells(newitemrow, methodcol).Value = "Extrapolated"
                            'Cells(newitemrow, lastareacol).Value = Cells(loc.Row, lastareacol).Value
                            'Cells(newitemrow, lastfaccol).Value = Cells(loc.Row, lastfaccol).Value
                            Set loc = .FindNext(loc)
                            foundnext = loc.Row
                            'newitemrow = newitemrow + 1
                        Else
                            finalformula = Replace(draftformula, "[rowholder]", loc.Row)
                            .Range("B" & newitemrow).Value = Range("B" & loc.Row).Value
                            .Range("C" & newitemrow).Value = Range("C" & loc.Row).Value
                            .Range("F" & newitemrow).Value = "=IFNA(VLOOKUP(""" & newcostcode & """, 'Cost Report'!G:I,3,FALSE),""NOT ASSIGNED"")"
                            .Range("G" & newitemrow).Value = newformulaname
                            .Range("H" & newitemrow).Value = "Extrapolated from " & Range("H" & loc.Row).Value
                            .Range("I" & newitemrow).Value = Range("I" & loc.Row).Value & "-E"
                            .Range("J" & newitemrow).Value = finalformula
                            .Range("K" & newitemrow).Value = newuom
                            With Cells(newitemrow, assemblycodecol)
                                .NumberFormat = "@"
                                .Value = newcostcode
                            End With
                            .Cells(newitemrow, milestonecol).Value = Cells(loc.Row, milestonecol).Value
                            .Cells(newitemrow, methodcol).Value = "Extrapolated"
                            .Cells(newitemrow, lastareacol).Value = Cells(loc.Row, lastareacol).Value
                            .Cells(newitemrow, lastfaccol).Value = Cells(loc.Row, lastfaccol).Value
                            Set loc = .FindNext(loc)
                            foundnext = loc.Row
                            newitemrow = newitemrow + 1
                        End If
                    Loop
                End If
            End With
            Set loc = Nothing
            With Range("A4:AO" & newitemrow - 1).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End With
        'Debug.Print "finished " & rule
    Next
    'finishtime = Time
    'Debug.Print finishtime - starttime
    Application.ScreenUpdating = True
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'REMOVE GROUPING ROWS FROM THE MASTER QTO REPORT
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub removegroupbys(firstposition)
    Dim itemcolumn As Long
    Dim lastrow As Long
    Dim firstitem As String
    Dim lastitem As String
    Dim groupbycount As Integer
    Dim columnname1 As String
    Dim firstgb As String
    Dim lastgb As String
    Dim gbaddr As String
    Dim itemaddr As String
    Dim x As Integer
    Dim gb As Variant
    Dim item As Variant
    Dim takeofftypecol As Integer
    Dim takeofftype As String

    With Worksheets("MasterQTO_flat")
        itemcolumn = Application.Match("Item", Sheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns("Column Name").DataBodyRange, False)
        lastrow = .Cells(Rows.count, itemcolumn).End(xlUp).Row
        firstitem = .Cells(firstposition, itemcolumn).Address
        lastitem = .Cells(lastrow, itemcolumn).Address
        groupbycount = itemcolumn - 2
        takeofftypecol = Application.Match("TE_QTO_Method", Sheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns("Column Name").DataBodyRange, False)
    
        If Not groupbycount = 0 Then
            For x = 2 To groupbycount + 1
                columnname1 = Worksheets("MasterQTO_flat").Cells(3, x).Value
                firstgb = Worksheets("MasterQTO_flat").Cells(firstposition, x).Address
                lastgb = Worksheets("MasterQTO_flat").Cells(lastrow, x).Address
                    For Each gb In Sheets("MasterQTO_flat").Range(firstgb & ":" & lastgb)
                        gbaddr = gb.Address
                        takeofftype = Worksheets("MasterQTO_flat").Cells(gb.Row, takeofftypecol).Value
                        If takeofftype <> "Manual" Then
                            If takeofftype <> "Extrapolated" Then
                                If Sheets("MasterQTO_flat").Range(gbaddr).Font.color = 0 Then
                                    Sheets("MasterQTO_flat").Range(gbaddr).EntireRow.Delete
                                End If
                            End If
                        End If
                    Next
            Next x
        End If
    
        For Each item In Sheets("MasterQTO_flat").Range(firstitem & ":" & lastitem)
            itemaddr = item.Address
            takeofftype = Worksheets("MasterQTO_flat").Cells(item.Row, takeofftypecol).Value
            If takeofftype <> "Manual" Then
                If takeofftype <> "Extrapolated" Then
                    If Sheets("MasterQTO_flat").Range(itemaddr).Font.color = 0 Then
                        Sheets("MasterQTO_flat").Range(itemaddr).EntireRow.Delete
                    End If
                End If
            End If
        Next
        With .Range("A4:" & lastitem)
            .Font.color = 0
            .Font.Bold = False
            .WrapText = False
            .EntireColumn.AutoFit
            .WrapText = True
        End With
    End With
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'RUN COST CODE COMPARISION REPORT FOR EACH FACILITY AND ALL FACILITIES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SummaryCOSTCODEComparison()
    Application.ScreenUpdating = False

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
    'Set ws_compare = Worksheets("Comparison")
    Set ws_QTOflat = Worksheets("MasterQTO_flat")
    'Set ws_trends = Worksheets("Trends")
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
    Dim trenddataExists As Boolean
    Dim trendTable As ListObject
    
    'VERIFY THAT THE MASTERQTO_FLAT SHEET EXISTS
    For Each sheet In wb_master.Sheets
        If sheet.Name = ws_QTOflat.Name Then
            QTO_flatExists = True
        End If
        If sheet.Name = "TrendData" Then
            trenddataExists = True
        End If
    Next
    If QTO_flatExists = False Then
        MsgBox "It looks like you have not imported QEX or MTO files. Please import data by using the 'Combine QEX Files' command and selecting QEX or MTO files."
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    'IF TRENDDATA DOES NOT EXIST, CREATE IT
    If trenddataExists = False Then
        Set ws_trenddata = Worksheets.Add(After:=Sheets(Sheets.count))
        With ws_trenddata
            .Name = "TrendData"
            .Visible = False
            Set trendTable = .ListObjects.Add(xlSrcRange)
            .Range("A1").Value = "Item"
            .Range("B1").Value = Date
        End With
    Else
        Set ws_trenddata = Worksheets("TrendData")
    End If
    

    
    'IDENTIFY WHICH COLUMNS TO USE FROM THE QTO_CONFIG TABLE
    areacol = Application.Match("TE_Area", lo_qto.DataBodyRange.Columns(2), False)
    facilitycol = Application.Match("TE_Facility", lo_qto.DataBodyRange.Columns(2), False)
    assemblycol = Application.Match("Assembly Code Level 3", lo_qto.DataBodyRange.Columns(2), False)
    costcodecol = Application.Match("Assembly Code", lo_qto.DataBodyRange.Columns(2), False)
    lastcol = lo_qto.DataBodyRange.Rows.count
    lastrow = ws_QTOflat.Cells(Rows.count, costcodecol).End(xlUp).Row
    lastitem = ws_QTOflat.Cells(lastrow, lastcol).Address
    
    'IDENTIFY LOCATION OF REQUIRED COLUMNS IN THE COST REPORT IN THE SITUATION THAT THEY ARE MOVED
    reportArea = Application.Match("Area", Sheets("Cost Report").Range("A1:BB1"), False)
        If reportArea <= 26 Then
            reportArea = Chr(reportArea + 64)
        Else
            lRemainder = reportArea Mod 26
            lAlpha = Int(reportArea / 26)
            If lRemainder = 0 Then
                lRemainder = 26
                lAlpha = lAlpha - 1
            End If
            reportArea = reportArea(lAlpha) & Chr(lRemainder + 64)
        End If
    reportFacility = Application.Match("Facility", Sheets("Cost Report").Range("A1:BB1"), False)
        If reportFacility <= 26 Then
            reportFacility = Chr(reportFacility + 64)
        Else
            lRemainder = reportFacility Mod 26
            lAlpha = Int(reportFacility / 26)
            If lRemainder = 0 Then
                lRemainder = 26
                lAlpha = lAlpha - 1
            End If
            reportFacility = reportFacility(lAlpha) & Chr(lRemainder + 64)
        End If
    reportFD = Application.Match("Facility Description", Sheets("Cost Report").Range("A1:BB1"), False)
        If reportFD <= 26 Then
            reportFD = Chr(reportFD + 64)
        Else
            lRemainder = reportFD Mod 26
            lAlpha = Int(reportFD / 26)
            If lRemainder = 0 Then
                lRemainder = 26
                lAlpha = lAlpha - 1
            End If
            reportFD = reportFD(lAlpha) & Chr(lRemainder + 64)
        End If
    reportCC = Application.Match("Cost Cde", Sheets("Cost Report").Range("A1:BB1"), False)
        If reportCC <= 26 Then
            reportCC = Chr(reportCC + 64)
        Else
            lRemainder = reportCC Mod 26
            lAlpha = Int(reportCC / 26)
            If lRemainder = 0 Then
                lRemainder = 26
                lAlpha = lAlpha - 1
            End If
            reportCC = reportCC(lAlpha) & Chr(lRemainder + 64)
        End If
    reportCCD = Application.Match("Cost Code Description", Sheets("Cost Report").Range("A1:BB1"), False)
        If reportCCD <= 26 Then
            reportCCD = Chr(reportCCD + 64)
        Else
            lRemainder = reportCCD Mod 26
            lAlpha = Int(reportCCD / 26)
            If lRemainder = 0 Then
                lRemainder = 26
                lAlpha = lAlpha - 1
            End If
            reportCCD = reportCCD(lAlpha) & Chr(lRemainder + 64)
        End If
    reportResource = Application.Match("Resource", Sheets("Cost Report").Range("A1:BB1"), False)
        If reportResource <= 26 Then
            reportResource = Chr(reportResource + 64)
        Else
            lRemainder = reportResource Mod 26
            lAlpha = Int(reportResource / 26)
            If lRemainder = 0 Then
                lRemainder = 26
                lAlpha = lAlpha - 1
            End If
            reportResource = reportResource(lAlpha) & Chr(lRemainder + 64)
        End If
    reportUOM = Application.Match("UOM", Sheets("Cost Report").Range("A1:BB1"), False)
        If reportUOM <= 26 Then
            reportUOM = Chr(reportUOM + 64)
        Else
            lRemainder = reportUOM Mod 26
            lAlpha = Int(reportUOM / 26)
            If lRemainder = 0 Then
                lRemainder = 26
                lAlpha = lAlpha - 1
            End If
            reportUOM = reportUOM(lAlpha) & Chr(lRemainder + 64)
        End If
    reportQTY = Application.Match("CUR QTY", Sheets("Cost Report").Range("A1:BB1"), False)
        If reportQTY <= 26 Then
            reportQTY = Chr(reportQTY + 64)
        Else
            lRemainder = reportQTY Mod 26
            lAlpha = Int(reportQTY / 26)
            If lRemainder = 0 Then
                lRemainder = 26
                lAlpha = lAlpha - 1
            End If
            reportQTY = reportArea(lAlpha) & Chr(lRemainder + 64)
        End If
    reportcombined = Application.Match("Area-Fac-Code-Resource-UOM", Sheets("Cost Report").Range("A1:BB1"), False)
    lastreportcode = Worksheets("Cost Report").Cells(Rows.count, 1).End(xlUp).Row
        If IsError(reportcombined) Then
            'reportLastcol
            Sheets("Cost Report").Range("AK1").Value = "Area-Fac-Code-Resource-UOM"
            With Sheets("Cost Report").Range("AK2:AK" & lastreportcode)
                .Value = "=CONCATENATE(" & reportArea & "2, ""-""," & reportFacility & "2, ""-""," & reportCC & "2, ""-""," & reportResource & "2, ""-""," & reportUOM & "2)"
                .FormatConditions.Delete
                .FormatConditions.AddUniqueValues
                .FormatConditions(1).DupeUnique = xlDuplicate
                .FormatConditions(1).Interior.color = 13551615
                .FormatConditions.Add Type:=xlExpression, Formula1:="=$" & reportQTY & "2=0"
                .FormatConditions(2).Interior.color = 49407
            End With
            reportcombined = "AK"
            
        End If
            
    'IF COMPARISON SHEET EXISTS, DELETE IT
    For Each sheet In ActiveWorkbook.Sheets
        If sheet.Name = "Comparison" Then
            Application.DisplayAlerts = False
            Worksheets("Comparison").Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    'SET QTO_CONFIG TABLE AS ARRAY TO VERIFY THE HEADER TRUE/FALSE VALUES FOR QTO Summary
    CompareConfig = Sheets("CONFIG").ListObjects("COMPARE_CONFIG").DataBodyRange

    isdimensioned = False

    'CHECK THE MANUAL COLUMN AND IF IT IS TRUE, ADD THE COLUMN NAME TO COLUMNHEADERS
    For x = LBound(CompareConfig) To UBound(CompareConfig)
            If isdimensioned = True Then
                ReDim Preserve columnheaders(1 To UBound(columnheaders) + 1) As String
            Else
                ReDim columnheaders(1 To 1) As String
                isdimensioned = True
            End If
            columnheaders(UBound(columnheaders)) = CompareConfig(x, 2)
    Next
            
    Worksheets.Add(After:=Sheets("MasterQTO_flat")).Name = "Comparison"
    Set ws_compare = Worksheets("Comparison")
        
    'SET UP THE HEADERS FOR ALL COST CODES
    With Sheets("Comparison")
        With .Range("A1")
            .Value = "EVLLRT-All Areas Cost Code Comparison Report"
            .Font.Bold = True
            .Font.Size = 14
        End With
            
        'GET THE ADDRESS OF THE LAST COLUMN HEADER NEEDED
        lastcol = UBound(columnheaders)
        lastheader = Cells(3, lastcol).Address
        With .Range("A3:" & lastheader)
            .Value = columnheaders
            .EntireColumn.AutoFit
            .Font.color = 16777215
            .Font.Bold = True
            .Interior.color = 5855577
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End With
    End With
    
    reportnotescol = Application.Match("Cost Report Notes", Worksheets("CONFIG").ListObjects("COMPARE_CONFIG").DataBodyRange.Columns(2), False)
    
    itemrow = 4
    
    Sheets("MasterQTO_flat").Activate
    
    
    
    
    
    'NEED TO COME BACK TO HERE: PROVIDE THE ABILITY TO SEARCH THE TRENDDATA SHEET FOR COSTCODE:AREA:FACILITY
    'AND IF FOUND, PLACE THE NEW VALUES FOR MODEL AND ESTIMATE. IF NOT FOUND, ADD NEW LINE WITH ITEM NAME, MODEL AND ESTIMATE QANTITY
    
    
    
    
    
    
    
    'SORT THE FLAT QTO BY AREA>FACILITY>ASSEMBLY CODE>ITEM
    With Sheets("MasterQTO_flat")
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=Cells(3, costcodecol)
        .Sort.SortFields.Add Key:=Cells(3, areacol)
        .Sort.SortFields.Add Key:=Cells(3, facilitycol)
        '.Sort.SortFields.Add Key:=Cells(3, costcodecol)
        .Sort.SetRange Range("A3:" & lastitem)
        .Sort.Header = xlYes
        .Sort.Apply
    End With
     
    'LOOP THROUGH EACH COST CODE, AREA, AND FACILITY TO CAPTURE THE VALUES AT EACH LEVEL'
    With ws_QTOflat
        'RUN THROUGH EACH COST CODE'
        For Each item In Range(Cells(4, costcodecol), Cells(lastrow, costcodecol))
            If item <> pItem And item <> "" Then
                With Sheets("Comparison")
                    .Cells(itemrow, 1).Value = "Cost Code " & item
                    .Cells(itemrow, 2).Value = "=VLOOKUP(""" & item & """,'Assembly Codes & Unit Costs'!A:G,2,FALSE)"
                    
                    '.Cells(itemrow, 1).Font.Bold = True
                    With .Range("A" & itemrow & ":" & Cells(itemrow, lastcol).Address)
                        .Interior.color = 8421504
                        .Font.Bold = True
                        '.Font.color = 16777215
                    End With
                End With
                
                pItem = item
                pItemRow = itemrow
                itemrow = itemrow + 1

                'RUN THROUGH EACH AREA
                For Each area In Range(.Cells(4, areacol), .Cells(lastrow, areacol))
                    If area <> pArea And area <> "Not Assigned" And .Cells(area.Row, costcodecol).Value = item Then
                        With Sheets("Comparison").Cells(itemrow, 1) 'ws_compare.Cells(itemrow, 1)
                            
                            .Value = "Area " & area
                            .Font.Bold = True
                            .IndentLevel = 1
                        End With
                            '.Cells(itemrow, 2).Value = "=VLOOKUP(""" & facility & """,'Cost Report'!" & reportFacility & ":" & reportFD & ",2,FALSE)"
                        ws_compare.Range("A" & itemrow & ":" & Cells(itemrow, lastcol).Address).Interior.color = 10921638
                        
                        pArea = area
                        pAreaRow = itemrow
                        itemrow = itemrow + 1

                        'RUN THROUGH EACH FACILITY
                        For Each facility In Range(Cells(4, facilitycol), Cells(lastrow, facilitycol))
                            If facility <> pFacility And facility <> "Not Assigned" And Cells(facility.Row, costcodecol).Value = item And Cells(facility.Row, areacol).Value = area Then
                                With ws_compare
                                    .Cells(itemrow, 1).Value = "Facility " & facility
                                    .Cells(itemrow, 1).IndentLevel = 3

                                    costcodecol = Application.Match("Assembly Code", Worksheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Columns(2), False)
                                    costcode = Sheets("MasterQTO_flat").Cells(item.Row, costcodecol).Value
                                    
                                    'TEMP TO ADD THE DATA TO TRENDS SHEET
                                    'Sheets("trends").Range("X" & itemrow).Value = area & "-" & facility & "-" & costcode
                                    For y = 2 To UBound(columnheaders)
                                        itemvalue = columnheaders(y)
                                        If itemvalue <> "Item" Then
                                            If itemvalue = "Facility Description" Then
                                                .Cells(itemrow, y).Value = "=VLOOKUP(""" & facility & """,'Cost Report'!" & reportFacility & ":" & reportFD & ",2,FALSE)"
                                            End If
                                            If itemvalue = "Cost Code Description" Then
                                                .Cells(itemrow, y).Value = "=VLOOKUP(""" & costcode & """,'Assembly Codes & Unit Costs'!A:G,2,FALSE)"
                                            End If
                                            If itemvalue = "UOM" Then
                                                itemvaluecol = Application.Match("Unit", Worksheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Columns(2), False)
                                                currentUOM = Sheets("MasterQTO_flat").Cells(item.Row, itemvaluecol).Value
                                                'SET THE MODEL UOM TO MATCH METRIC TON
                                                If currentUOM = "T" Then
                                                    currentUOM = "MT"
                                                End If
                                                .Cells(itemrow, y).Value = currentUOM

                                            End If
                                            If itemvalue = "Model Milestone" Then
                                                itemvaluecol = Application.Match("TE_Milestone", Worksheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Columns(2), False)
                                                .Cells(itemrow, y).Value = Sheets("MasterQTO_flat").Cells(item.Row, itemvaluecol).Value
                                            End If
                                            If itemvalue = "QTO Source" Then
                                                itemvaluecol = Application.Match("TE_QTO_Source", Worksheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Columns(2), False)
                                                .Cells(itemrow, y).Value = Sheets("MasterQTO_flat").Cells(item.Row, itemvaluecol).Value
                                            End If
                                            If itemvalue = "QTO Method" Then
                                                itemvaluecol = Application.Match("TE_QTO_Method", Worksheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Columns(2), False)
                                                .Cells(itemrow, y).Value = Sheets("MasterQTO_flat").Cells(item.Row, itemvaluecol).Value
                                            End If
                                            If itemvalue = "Model QTY" Then
                                                'IF THERE ARE QUOTES IN THE ASSEMBLY CODE OR DESCRIPTION, REPLACE THEM WITH DOUBLE QUOTES TO BE PLACED IN FORMULAS
                                                assemblycodeformula = Replace(assemblycode, Chr(34), Chr(34) & Chr(34))
                                                itemformula = Replace(item, Chr(34), Chr(34) & Chr(34))

                                                itemvaluecol = Application.Match("Quantity", Worksheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Columns(2), False)
                                                .Cells(itemrow, y).Value = "=SumIfs(MasterQTO_flat!" & .Cells(4, itemvaluecol).Address & ":" & .Cells(lastrow, itemvaluecol).Address & ", MasterQTO_flat!" & .Cells(4, areacol).Address & ":" & .Cells(lastrow, areacol).Address & ", """ & area & """, MasterQTO_flat!" & .Cells(4, facilitycol).Address & ":" & .Cells(lastrow, facilitycol).Address & ", """ & facility & """, MasterQTO_flat!" & .Cells(4, costcodecol).Address & ":" & .Cells(lastrow, costcodecol).Address & ", """ & itemformula & """)"
                                                QTOflat_qcol = itemvaluecol
                                                'THE NEXT ROW CONTAINS THE ASSEMBLY CODE VALUE AS WELL AS THE COST CODE
                                                '.Cells(itemrow, y).Value = "=SumIfs(MasterQTO_flat!" & .Cells(4, itemvaluecol).Address & ":" & .Cells(lastrow, itemvaluecol).Address & ", MasterQTO_flat!" & .Cells(4, areacol).Address & ":" & .Cells(lastrow, areacol).Address & ", """ & area & """, MasterQTO_flat!" & .Cells(4, facilitycol).Address & ":" & .Cells(lastrow, facilitycol).Address & ", """ & facility & """, MasterQTO_flat!" & .Cells(4, assemblycol).Address & ":" & .Cells(lastrow, assemblycol).Address & ", """ & assemblycodeformula & """,MasterQTO_flat!" & .Cells(4, costcodecol).Address & ":" & .Cells(lastrow, costcodecol).Address & ", """ & itemformula & """)"
                                                .Cells(itemrow, y).NumberFormat = 0#
                                            End If
                                            If itemvalue = "Estimate QTY" Then
                                                .Cells(itemrow, y).Value = "No Estimate"
                                                .Cells(itemrow, y).Interior.color = 13551615
                                                .Cells(itemrow, y).Font.color = 393372
                                                '.Cells(itemrow, y).Style = "Bad"
                                                MATLcode = area & "-" & facility & "-" & costcode & "-MATL-" & currentUOM
                                                MLcode = area & "-" & facility & "-" & costcode & "-ML-" & currentUOM
                                                SCcode = area & "-" & facility & "-" & costcode & "-SC-" & currentUOM

                                            'SEARCH FOR THE MATL CODE AND IF FAILS, SEARCH FOR THE ML then SC. IF FOUND AND RED, SHOW AS MULTIPLE VALUES. ELSE GRAB TAKEOFF VALUE
                                            searchnext = True
                                            coderow = Application.Match(MATLcode, Sheets("Cost Report").Range("AK:AK"), False)
                                            If Not IsError(coderow) Then
                                                If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 16777215 Then
                                                    .Cells(itemrow, y).Value = Sheets("Cost Report").Range(reportQTY & coderow).Value
                                                    .Cells(itemrow, y).Interior.color = xlNone
                                                    .Cells(itemrow, y).Font.color = 0
                                                    .Cells(itemrow, reportnotescol).Value = "MATL found on row " & coderow
                                                    searchnext = False
                                                Else
                                                    If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 13551615 Then
                                                        .Cells(itemrow, y).Value = "Multiple Estimates"
                                                        .Cells(itemrow, y).Interior.color = 13551615
                                                        .Cells(itemrow, y).Font.color = 393372
                                                    End If
                                                    If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 49407 Then
                                                        .Cells(itemrow, y).Value = "0"
                                                        .Cells(itemrow, y).Interior.color = xlNone
                                                        .Cells(itemrow, y).Font.color = 0
                                                        .Cells(itemrow, reportnotescol).Value = "MATL found on row " & coderow
                                                    End If
                                                End If
                                            End If
                                            If searchnext = True Then
                                                coderow = Application.Match(MLcode, Sheets("Cost Report").Range("AK:AK"), False)
                                                If Not IsError(coderow) Then
                                                    If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 16777215 Then
                                                        .Cells(itemrow, y).Value = Sheets("Cost Report").Range(reportQTY & coderow).Value
                                                        .Cells(itemrow, reportnotescol).Value = "ML found on row " & coderow
                                                        .Cells(itemrow, y).Interior.color = xlNone
                                                        .Cells(itemrow, y).Font.color = 0
                                                        searchnext = False
                                                    Else
                                                        If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 13551615 Then
                                                            .Cells(itemrow, y).Value = "Multiple Estimates"
                                                            .Cells(itemrow, y).Interior.color = 13551615
                                                            .Cells(itemrow, y).Font.color = 393372
                                                        End If
                                                        If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 49407 Then
                                                            .Cells(itemrow, y).Value = "0"
                                                            .Cells(itemrow, y).Interior.color = xlNone
                                                            .Cells(itemrow, y).Font.color = 0
                                                            .Cells(itemrow, reportnotescol).Value = "MATL found on row " & coderow
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            If searchnext = True Then
                                                coderow = Application.Match(SCcode, Sheets("Cost Report").Range("AK:AK"), False)
                                                If Not IsError(coderow) Then
                                                    If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 16777215 Then
                                                        .Cells(itemrow, y).Value = Sheets("Cost Report").Range(reportQTY & coderow).Value
                                                        .Cells(itemrow, reportnotescol).Value = "SC found on row " & coderow
                                                        .Cells(itemrow, y).Interior.color = xlNone
                                                        .Cells(itemrow, y).Font.color = 0
                                                    Else
                                                        If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 13551615 Then
                                                            .Cells(itemrow, y).Value = "Multiple Estimates"
                                                            .Cells(itemrow, y).Interior.color = 13551615
                                                            .Cells(itemrow, y).Font.color = 393372
                                                        End If
                                                        If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 49407 Then
                                                            .Cells(itemrow, y).Value = "0"
                                                            .Cells(itemrow, y).Interior.color = xlNone
                                                            .Cells(itemrow, y).Font.color = 0
                                                            .Cells(itemrow, reportnotescol).Value = "MATL found on row " & coderow
                                                        End If
                                                    End If
                                                End If
                                            End If

                                                '.Cells(itemrow, y).Value = "=SumIfs('Cost Report'!O2:O2877, 'Cost Report'!B2:B2877, """ & area & """, 'Cost Report'!C2:C2877, """ & facility & """, 'Cost Report'!G2:G2877, """ & costcode & """, 'Cost Report'!L2:L2877, ""MATL"")"
                                                .Cells(itemrow, y).NumberFormat = 0#
                                            End If
                                            If itemvalue = "Estimate QTY Difference" Then
                                                'REMOVE STATIC REFERENCE VALUES
                                                .Cells(itemrow, y).Value = "=IF(F" & itemrow & "=""No QTY"",-G" & itemrow & ",IF(OR(G" & itemrow & "=""No Estimate"",G" & itemrow & "=""Multiple Estimates""),F" & itemrow & ",F" & itemrow & "-G" & itemrow & "))"
                                                .Cells(itemrow, y).NumberFormat = 0#
                                            End If
                                            If itemvalue = "Estimate % Deviation" Then
                                                .Cells(itemrow, y).Value = "=IF(F" & itemrow & "=G" & itemrow & ",0,IF(OR(F" & itemrow & "=0,F" & itemrow & "=""No QTY""),-1,IF(OR(G" & itemrow & "=0,G" & itemrow & "=""No Estimate"",G" & itemrow & "=""Multiple Estimates""),1,IF(F" & itemrow & ">G" & itemrow & ",(F" & itemrow & "/G" & itemrow & ")-1,-((G" & itemrow & "/F" & itemrow & ")-1)))))"
                                                .Cells(itemrow, y).NumberFormat = "0.00%"
                                            End If
                                            If itemvalue = "Within Tolerance" Then
                                                .Cells(itemrow, y).Value = "=IFERROR(IF(ABS(I" & itemrow & ")>VLOOKUP(E" & itemrow & ",Tolerance_CONFIG,2,FALSE),""No"",""Yes""),""No"")"
                                                    .Cells.FormatConditions.Delete
                                                    With Sheets("Comparison").Range("J:J").FormatConditions _
                                                        .Add(xlCellValue, xlEqual, "Yes")
                                                        '.Style = "Good"
                                                        .Interior.color = 13561798
                                                        .Font.color = 24832
                                                    End With
                                                    With Sheets("Comparison").Range("J:J").FormatConditions _
                                                        .Add(xlCellValue, xlEqual, "No")
                                                        '.Style = "Bad"
                                                        .Interior.color = 13551615
                                                        .Font.color = 393372
                                                    End With

                                            End If
                                            If itemvalue = "Responsible Function" Then
                                                .Cells(itemrow, y).Value = "=VLOOKUP(""" & item & """,'Assembly Codes & Unit Costs'!A:H,8,FALSE)"
                                            End If
                                            If itemvalue = "Action Owner" Then
                                                .Cells(itemrow, y).Value = ""
                                            End If
                                            If itemvalue = "Action" Then
                                                .Cells(itemrow, y).Value = ""
                                            End If
                                            If itemvalue = "Area" Then
                                                .Cells(itemrow, y).Value = area
                                            End If
                                        End If
                                    Next
                                End With
                                pFacility = facility
                                pFacilityRow = itemrow
                                itemrow = itemrow + 1
                            End If
                        Next
                        'INCORPORATE THE COSTCODES FROM THE COST REPORT THAT DO NOT HAVE A MODELED COMPONENT
                        With Sheets("Cost Report")
                            For Each facility In .Range(reportFacility & "1:" & reportFacility & lastreportcode)
                                'CHECK TO SEE IF THE REPORT CODE IS WITHIN THE 89 CODES
                                reportcode = .Range(reportCC & facility.Row).Value
                                foundinreport = Application.Match(reportcode, Worksheets("Assembly Codes & Unit Costs").Range("A1:A90"), False)
                                If Not IsError(foundinreport) Then
                                    'CHECK TO SEE IF COMBINED VALUE IS A MATCH
                                    uom = Sheets("Assembly Codes & Unit Costs").Cells(foundinreport, 5).Value
                                    combinedvalfound = False
                                    If .Cells(facility.Row, 37).Value = area & "-" & facility & "-" & item & "-" & "MATL" & "-" & uom Then combinedvalfound = True
                                    If .Cells(facility.Row, 37).Value = area & "-" & facility & "-" & item & "-" & "ML" & "-" & uom Then combinedvalfound = True
                                    If .Cells(facility.Row, 37).Value = area & "-" & facility & "-" & item & "-" & "SC" & "-" & uom Then combinedvalfound = True
                                    If combinedvalfound = True Then
                                        alreadyincomparison = Application.Match("Facility " & facility, Worksheets("Comparison").Range("A" & pAreaRow + 1 & ":" & "A" & itemrow), False)
                                        If IsError(alreadyincomparison) Then
                                            With Sheets("Comparison")
                                                .Cells(itemrow, 1).Value = "Facility " & facility
                                                .Cells(itemrow, 1).IndentLevel = 3
                                                For y = 2 To UBound(columnheaders)
                                                    itemvalue = columnheaders(y)
                                                    If itemvalue <> "Item" Then
                                                        If itemvalue = "Facility Description" Then
                                                            .Cells(itemrow, y).Value = "=VLOOKUP(""" & facility & """,'Cost Report'!" & reportFacility & ":" & reportFD & ",2,FALSE)"
                                                        End If
                                                        If itemvalue = "Cost Code Description" Then
                                                            .Cells(itemrow, y).Value = "=VLOOKUP(""" & reportcode & """,'Assembly Codes & Unit Costs'!A:G,2,FALSE)"
                                                        End If
                                                        If itemvalue = "UOM" Then
                                                            .Cells(itemrow, y).Value = uom
                                                        End If
                                                        If itemvalue = "Model Milestone" Then
                                                            .Cells(itemrow, y).Value = "No Model"
                                                            .Cells(itemrow, y).Interior.color = 13551615
                                                            .Cells(itemrow, y).Font.color = 393372
                                                        End If
                                                        If itemvalue = "Model QTY" Then
                                                            .Cells(itemrow, y).Value = "No QTY"
                                                            .Cells(itemrow, y).Interior.color = 13551615
                                                            .Cells(itemrow, y).Font.color = 393372
                                                        End If
                                                        If itemvalue = "Estimate QTY" Then
                                                            .Cells(itemrow, y).Value = "No Estimate"
                                                            .Cells(itemrow, y).Interior.color = 13551615
                                                            .Cells(itemrow, y).Font.color = 393372
                                                            '.Cells(itemrow, y).Style = "Bad"
                                                            MATLcode = area & "-" & facility & "-" & reportcode & "-MATL-" & uom
                                                            MLcode = area & "-" & facility & "-" & reportcode & "-ML-" & uom
                                                            SCcode = area & "-" & facility & "-" & reportcode & "-SC-" & uom

                                                            'SEARCH FOR THE MATL CODE AND IF FAILS, SEARCH FOR THE ML then SC. IF FOUND AND RED, SHOW AS MULTIPLE VALUES. ELSE GRAB TAKEOFF VALUE
                                                            searchnext = True
                                                            coderow = Application.Match(MATLcode, Sheets("Cost Report").Range("AK:AK"), False)
                                                            If Not IsError(coderow) Then
                                                                If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 16777215 Then
                                                                    .Cells(itemrow, y).Value = Sheets("Cost Report").Range(reportQTY & coderow).Value
                                                                    .Cells(itemrow, y).Interior.color = xlNone
                                                                    .Cells(itemrow, y).Font.color = 0
                                                                    .Cells(itemrow, reportnotescol).Value = "MATL found on row " & coderow
                                                                    searchnext = False
                                                                Else
                                                                    If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 13551615 Then
                                                                        .Cells(itemrow, y).Value = "Multiple Estimates"
                                                                        .Cells(itemrow, y).Interior.color = 13551615
                                                                        .Cells(itemrow, y).Font.color = 393372
                                                                    End If
                                                                    If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 49407 Then
                                                                        .Cells(itemrow, y).Value = "0"
                                                                        .Cells(itemrow, y).Interior.color = xlNone
                                                                        .Cells(itemrow, y).Font.color = 0
                                                                        .Cells(itemrow, reportnotescol).Value = "MATL found on row " & coderow
                                                                    End If
                                                                End If
                                                            End If
                                                            If searchnext = True Then
                                                                coderow = Application.Match(MLcode, Sheets("Cost Report").Range("AK:AK"), False)
                                                                If Not IsError(coderow) Then
                                                                    If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 16777215 Then
                                                                        .Cells(itemrow, y).Value = Sheets("Cost Report").Range(reportQTY & coderow).Value
                                                                        .Cells(itemrow, reportnotescol).Value = "ML found on row " & coderow
                                                                        .Cells(itemrow, y).Interior.color = xlNone
                                                                        .Cells(itemrow, y).Font.color = 0
                                                                        searchnext = False
                                                                    Else
                                                                        If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 13551615 Then
                                                                            .Cells(itemrow, y).Value = "Multiple Estimates"
                                                                            .Cells(itemrow, y).Interior.color = 13551615
                                                                            .Cells(itemrow, y).Font.color = 393372
                                                                        End If
                                                                        If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 49407 Then
                                                                            .Cells(itemrow, y).Value = "0"
                                                                            .Cells(itemrow, y).Interior.color = xlNone
                                                                            .Cells(itemrow, y).Font.color = 0
                                                                            .Cells(itemrow, reportnotescol).Value = "MATL found on row " & coderow
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                            If searchnext = True Then
                                                                coderow = Application.Match(SCcode, Sheets("Cost Report").Range("AK:AK"), False)
                                                                If Not IsError(coderow) Then
                                                                    If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 16777215 Then
                                                                        .Cells(itemrow, y).Value = Sheets("Cost Report").Range(reportQTY & coderow).Value
                                                                        .Cells(itemrow, reportnotescol).Value = "SC found on row " & coderow
                                                                        .Cells(itemrow, y).Interior.color = xlNone
                                                                        .Cells(itemrow, y).Font.color = 0
                                                                    Else
                                                                        If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 13551615 Then
                                                                            .Cells(itemrow, y).Value = "Multiple Estimates"
                                                                            .Cells(itemrow, y).Interior.color = 13551615
                                                                            .Cells(itemrow, y).Font.color = 393372
                                                                        End If
                                                                        If Sheets("Cost Report").Range("AK" & coderow).DisplayFormat.Interior.color = 49407 Then
                                                                            .Cells(itemrow, y).Value = "0"
                                                                            .Cells(itemrow, y).Interior.color = xlNone
                                                                            .Cells(itemrow, y).Font.color = 0
                                                                            .Cells(itemrow, reportnotescol).Value = "MATL found on row " & coderow
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                            .Cells(itemrow, y).NumberFormat = 0#
                                                        End If
                                                        If itemvalue = "Estimate QTY Difference" Then
                                                            'REMOVE STATIC REFERENCE VALUES
                                                            .Cells(itemrow, y).Value = "=IFERROR(IF(F" & itemrow & "=""No QTY"",-G" & itemrow & ",IF(OR(G" & itemrow & "=""No Estimate"",G" & itemrow & "=""Multiple Estimates""),F" & itemrow & ",F" & itemrow & "-G" & itemrow & ")),0)"
                                                            .Cells(itemrow, y).NumberFormat = 0#
                                                        End If
                                                        If itemvalue = "Estimate % Deviation" Then
                                                            .Cells(itemrow, y).Value = "=IF(F" & itemrow & "=G" & itemrow & ",0,IF(OR(F" & itemrow & "=0,F" & itemrow & "=""No QTY""),-1,IF(OR(G" & itemrow & "=0,G" & itemrow & "=""No Estimate"",G" & itemrow & "=""Multiple Estimates""),1,IF(F" & itemrow & ">G" & itemrow & ",(F" & itemrow & "/G" & itemrow & ")-1,-((G" & itemrow & "/F" & itemrow & ")-1)))))"
                                                            .Cells(itemrow, y).NumberFormat = "0.00%"
                                                        End If
                                                        If itemvalue = "Within Tolerance" Then
                                                            .Cells(itemrow, y).Value = "=IFERROR(IF(ABS(I" & itemrow & ")>VLOOKUP(E" & itemrow & ",Tolerance_CONFIG,2,FALSE),""No"",""Yes""),""No"")"
                                                                .Cells.FormatConditions.Delete
                                                                With Sheets("Comparison").Range("J:J").FormatConditions _
                                                                    .Add(xlCellValue, xlEqual, "Yes")
                                                                    '.Style = "Good"
                                                                    .Interior.color = 13561798
                                                                    .Font.color = 24832
                                                                End With
                                                                With Sheets("Comparison").Range("J:J").FormatConditions _
                                                                    .Add(xlCellValue, xlEqual, "No")
                                                                    '.Style = "Bad"
                                                                    .Interior.color = 13551615
                                                                    .Font.color = 393372
                                                                End With

                                                        End If
                                                        If itemvalue = "Responsible Function" Then
                                                            .Cells(itemrow, y).Value = "=VLOOKUP(""" & item & """,'Assembly Codes & Unit Costs'!A:H,8,FALSE)"
                                                        End If
                                                        If itemvalue = "Action Owner" Then
                                                            .Cells(itemrow, y).Value = ""
                                                        End If
                                                        If itemvalue = "Action" Then
                                                            .Cells(itemrow, y).Value = ""
                                                        End If
                                                        If itemvalue = "Area" Then
                                                            .Cells(itemrow, y).Value = area
                                                        End If
                                                    End If
                                                Next
                                            End With
                                            itemrow = itemrow + 1
                                        End If
                                    End If
                                End If
                            Next
                        End With
                        With ws_compare
                            For y = 2 To UBound(columnheaders)
                                itemvalue = columnheaders(y)
                                If itemvalue = "Cost Code Description" Then
                                    .Cells(pAreaRow, y).Value = "=VLOOKUP(""" & costcode & """,'Assembly Codes & Unit Costs'!A:G,2,FALSE)"
                                End If
                                If itemvalue = "UOM" Or itemvalue = "Model Milestone" Or itemvalue = "QTO Source" Or itemvalue = "QTO Method" Then 'THIS FORMULA NEEDS VALIDATION
                                    .Cells(pAreaRow, y).Value = "=IF(COUNTIF(" & .Cells(pAreaRow + 1, y).Address & ":" & .Cells(itemrow - 1, y).Address & "," & .Cells(pAreaRow + 1, y).Address & ")=COUNTA(" & .Cells(pAreaRow + 1, y).Address & ":" & .Cells(itemrow - 1, y).Address & ")=TRUE," & .Cells(pAreaRow + 1, y).Address & ","""")"
                                End If
                                If itemvalue = "Model QTY" Or itemvalue = "Estimate QTY" Or itemvalue = "Estimate QTY Difference" Then 'NEED TO ADJUST THIS ONE TO EXCLUDE THE INDIVIDUAL FACILITY AND ONLY SUM FOR CC AND AREA, OR USE A SUM VALUE
                                    .Cells(pAreaRow, y).Value = "=SUBTOTAL(9," & .Cells(pAreaRow + 1, y).Address & ":" & .Cells(itemrow - 1, y).Address & ")"
                                    .Cells(pAreaRow, y).NumberFormat = 0#
                                End If
                                If itemvalue = "Estimate % Deviation" Then 'NEED NEW FORMULA HERE
                                    .Cells(pAreaRow, y).Value = "=IF(F" & pAreaRow & "=G" & pAreaRow & ",0,IF(OR(F" & pAreaRow & "=0,F" & pAreaRow & "=""No QTY""),-1,IF(OR(G" & pAreaRow & "=0,G" & pAreaRow & "=""No Estimate"",G" & pAreaRow & "=""Multiple Estimates""),1,IF(F" & pAreaRow & ">G" & pAreaRow & ",(F" & pAreaRow & "/G" & pAreaRow & ")-1,-((G" & pAreaRow & "/F" & pAreaRow & ")-1)))))"
                                    .Cells(pAreaRow, y).NumberFormat = "0.00%"
                                End If
                                If itemvalue = "Within Tolerance" Then
                                .Cells(pAreaRow, y).Value = "=IFERROR(IF(ABS(I" & pAreaRow & ")>VLOOKUP(""Area Roll Up"",Tolerance_CONFIG,2,FALSE),""No"",""Yes""),""No"")"
                                End If
                                If itemvalue = "Responsible Function" Then
                                    .Cells(pAreaRow, y).Value = "=VLOOKUP(""" & item & """,'Assembly Codes & Unit Costs'!A:H,8,FALSE)"
                                End If
                                If itemvalue = "Area" Then
                                    .Cells(pAreaRow, y).Value = area
                                End If
                            Next
                            .Sort.SortFields.Clear
                            .Sort.SortFields.Add Key:=Cells(1, 1)
                            .Sort.SetRange Range(Cells(pFacilityRow + 1, 1), Cells(itemrow - 1, lastcol))
                            .Sort.Header = xlNo
                            .Sort.Apply
                        End With
                        
                        'CLOSE OUT THE AREA GROUPING
                        If pAreaRow <> itemrow Then
                            With Sheets("Comparison")
                                .Range("A" & pAreaRow + 1 & ":" & "A" & itemrow - 1).EntireRow.Group
                                .Outline.SummaryRow = False
                            End With
                            pFacility = "" 'Set pFacility = Nothing
                        End If
                    End If
                Next
                With ws_compare
                    For y = 2 To UBound(columnheaders)
                        itemvalue = columnheaders(y)
                        If itemvalue = "Cost Code Description" Then
                            .Cells(pItemRow, y).Value = "=VLOOKUP(""" & costcode & """,'Assembly Codes & Unit Costs'!A:G,2,FALSE)"
                        End If
                        If itemvalue = "UOM" Or itemvalue = "Model Milestone" Or itemvalue = "QTO Source" Or itemvalue = "QTO Method" Then 'THIS FORMULA NEEDS VALIDATION
                            .Cells(pItemRow, y).Value = "=IF(COUNTIF(" & .Cells(pItemRow + 1, y).Address & ":" & .Cells(itemrow - 1, y).Address & "," & .Cells(pItemRow + 1, y).Address & ")=COUNTA(" & .Cells(pItemRow + 1, y).Address & ":" & .Cells(itemrow - 1, y).Address & ")=TRUE," & .Cells(pItemRow + 1, y).Address & ","""")"
                        End If
                        If itemvalue = "Model QTY" Or itemvalue = "Estimate QTY" Or itemvalue = "Estimate QTY Difference" Then 'NEED TO ADJUST THIS ONE TO EXCLUDE THE INDIVIDUAL FACILITY AND ONLY SUM FOR CC AND AREA, OR USE A SUM VALUE
                            .Cells(pItemRow, y).Value = "=SUBTOTAL(9," & .Cells(pItemRow + 1, y).Address & ":" & .Cells(itemrow - 1, y).Address & ")"
                            .Cells(pItemRow, y).NumberFormat = 0#
                        End If
                        If itemvalue = "Estimate % Deviation" Then 'NEED NEW FORMULA HERE
                            .Cells(pItemRow, y).Value = "=IF(F" & pItemRow & "=G" & pItemRow & ",0,IF(OR(F" & pItemRow & "=0,F" & pItemRow & "=""No QTY""),-1,IF(OR(G" & pItemRow & "=0,G" & pItemRow & "=""No Estimate"",G" & pItemRow & "=""Multiple Estimates""),1,IF(F" & pItemRow & ">G" & pItemRow & ",(F" & pItemRow & "/G" & pItemRow & ")-1,-((G" & pItemRow & "/F" & pItemRow & ")-1)))))"
                            .Cells(pItemRow, y).NumberFormat = "0.00%"
                        End If
                        If itemvalue = "Within Tolerance" Then
                        .Cells(pItemRow, y).Value = "=IFERROR(IF(ABS(I" & pItemRow & ")>VLOOKUP(""Overall Cost Code"",Tolerance_CONFIG,2,FALSE),""No"",""Yes""),""No"")"
                        End If
                        If itemvalue = "Responsible Function" Then
                            .Cells(pItemRow, y).Value = "=VLOOKUP(""" & item & """,'Assembly Codes & Unit Costs'!A:H,8,FALSE)"
                        End If
                    Next
                End With
                
                'CLOSE OUT THE COST CODE GROUP
                If pItemRow <> itemrow Then
                    With ws_compare
                        .Range("A" & pItemRow + 1 & ":" & "A" & itemrow - 1).EntireRow.Group
                        .Outline.SummaryRow = False
                    End With
                    pArea = "" 'Set pArea = Nothing
                End If
            End If
        Next
    End With
    

    With Sheets("Comparison")
        .UsedRange.EntireColumn.AutoFit
        '.Outline.ShowLevels rowlevels:=3
    End With

    'SET CONDITIONAL FORMATTING FOR % DEVIATION COLUMN
    With Sheets("Comparison").Range("I:I")
        Set deviationcolor = .FormatConditions.AddColorScale(ColorScaleType:=3)
        deviationcolor.ColorScaleCriteria(1).Type = xlConditionValueNumber
        deviationcolor.ColorScaleCriteria(1).Value = -1
        deviationcolor.ColorScaleCriteria(1).FormatColor.color = RGB(255, 0, 0)
        deviationcolor.ColorScaleCriteria(2).Type = xlConditionValueNumber
        deviationcolor.ColorScaleCriteria(2).Value = 0
        deviationcolor.ColorScaleCriteria(2).FormatColor.color = RGB(255, 255, 255)
        deviationcolor.ColorScaleCriteria(3).Type = xlConditionValueNumber
        deviationcolor.ColorScaleCriteria(3).Value = 1
        deviationcolor.ColorScaleCriteria(3).FormatColor.color = RGB(255, 0, 0)
    End With

    Sheets("Comparison").Activate
    Sheets("Comparison").Range("A1").Select

Application.ScreenUpdating = True

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CREATE TEMPLATE FOR MANUAL TAKEOFFS WITH THE COLUMN HEADERS SELECTED IN CONFIG
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub MTOtemplate()

Application.ScreenUpdating = False

Dim MTOheaderConfig As Variant
Dim isdimensioned As Boolean
Dim x As Integer
Dim columnheaders() As String
Dim lastcol As Integer
Dim lastheader As String
Dim TemplateBook As Workbook
Dim lastgrid As String
    
    'SET QTO_CONFIG TABLE AS ARRAY TO VERIFY THE HEADER TRUE/FALSE VALUES FOR MTO TEMPLATE
    MTOheaderConfig = Sheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange
    
    isdimensioned = False
            
    'CHECK THE MANUAL COLUMN AND IF IT IS TRUE, ADD THE COLUMN NAME TO COLUMNHEADERS
    For x = LBound(MTOheaderConfig) To UBound(MTOheaderConfig)
        If MTOheaderConfig(x, 5) = True Then
            If isdimensioned = True Then
                ReDim Preserve columnheaders(1 To UBound(columnheaders) + 1) As String
            Else
                ReDim columnheaders(1 To 1) As String
                isdimensioned = True
            End If
            columnheaders(UBound(columnheaders)) = MTOheaderConfig(x, 2)
        End If
    Next
    
    'CREATE THE TEMPLATE WORKBOOK AND ADD THE COLUMNHEADERS
    Set TemplateBook = Workbooks.Add
    With TemplateBook
        With Sheets("Sheet1")
            .Name = "MTO-Template"
            With .Range("A1")
                .Value = "EVLLRT - Manual Takeoffs"
                .Font.Bold = True
                .Font.Size = 14
            End With
            'GET THE ADDRESS OF THE LAST COLUMN HEADER NEEDED
            lastcol = UBound(columnheaders)
            lastheader = Cells(3, lastcol).Address
            lastgrid = Cells(1000, lastcol).Address
            With .Range("A3:" & lastheader)
                .Value = columnheaders
                .EntireColumn.AutoFit
                .Font.color = 16777215
                .Font.Bold = True
                .Interior.color = 12419407
            End With
            With .Range("A3:" & lastgrid)
                With .Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
            End With
        End With
    End With
Application.ScreenUpdating = True
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CREATE SUMMARY QTO REPORT WITH PROPER GROUPINGS FOR EACH AREA, FACI
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SummaryQTO()
    Application.ScreenUpdating = False

    Dim areacol As Integer
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
    Dim lastitem As String
    Dim itemrow As Variant
    Dim pArea As String
    Dim pFacility As String
    Dim pAssemblyCode As String
    Dim pItem As String
    Dim columnheaders() As String
    Dim QTO_flatExists As Boolean
    Dim QTOSummaryConfig As Variant
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
    Dim quantitycol As Integer
    Dim sumstart As String
    Dim sumend As String
    Dim assemblycodeformula As String
    Dim itemformula As String
    Dim itemlong As String
    Dim pitemlong As String
    'Dim trenddatarow As Long
    
  
    'VERIFY THAT THE MASTERQTO_FLAT SHEET EXISTS
    For Each sheet In ActiveWorkbook.Sheets
        If sheet.Name = "MasterQTO_flat" Then
            QTO_flatExists = True
        End If
    Next
    If QTO_flatExists = False Then
        MsgBox "It looks like you have not imported QEX or MTO files. Please import data by using the 'Combine QEX Files' command and selecting QEX or MTO files."
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    'IDENTIFY WHICH COLUMNS TO USE FROM THE QTO_CONFIG TABLE
    areacol = Application.Match("TE_Area", Worksheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Columns(2), False)
    facilitycol = Application.Match("TE_Facility", Worksheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Columns(2), False)
    assemblycol = Application.Match("Assembly Code Level 3", Worksheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Columns(2), False)
    itemcol = Application.Match("Item", Worksheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Columns(2), False)
    lastcol = Worksheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Rows.count
    lastrow = Worksheets("MasterQTO_flat").Cells(Rows.count, itemcol).End(xlUp).Row
    lastitem = Worksheets("MasterQTO_flat").Cells(lastrow, lastcol).Address
    
    For Each sheet In ActiveWorkbook.Sheets
        If sheet.Name = "MasterQTO" Then
            Application.DisplayAlerts = False
            Worksheets("MasterQTO").Delete
            Application.DisplayAlerts = True
        End If
    Next
    
    'SET QTO_CONFIG TABLE AS ARRAY TO VERIFY THE HEADER TRUE/FALSE VALUES FOR QTO Summary
    QTOSummaryConfig = Sheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange
    
    isdimensioned = False
            
    'CHECK THE MANUAL COLUMN AND IF IT IS TRUE, ADD THE COLUMN NAME TO COLUMNHEADERS
    For x = LBound(QTOSummaryConfig) To UBound(QTOSummaryConfig)
        If QTOSummaryConfig(x, 6) = True Then
            If isdimensioned = True Then
                ReDim Preserve columnheaders(1 To UBound(columnheaders) + 1) As String
            Else
                ReDim columnheaders(1 To 1) As String
                isdimensioned = True
            End If
            columnheaders(UBound(columnheaders)) = QTOSummaryConfig(x, 2)
        End If
    Next
    
    'ADD THE RESPONSIBLE FUNCTION AT THE END
    ReDim Preserve columnheaders(1 To UBound(columnheaders) + 1) As String
    columnheaders(UBound(columnheaders)) = "Responsible Function"
            
    Worksheets.Add(After:=Sheets("MasterQTO_flat")).Name = "MasterQTO"
        
    With Sheets("MasterQTO")
        With .Range("A1")
            .Value = "EVLLRT-All Areas QTO Summary Report"
            .Font.Bold = True
            .Font.Size = 14
        End With
        
            
        'GET THE ADDRESS OF THE LAST COLUMN HEADER NEEDED
        lastcol = UBound(columnheaders)
        lastheader = Cells(3, lastcol).Address
        With .Range("A3:" & lastheader)
            .Value = columnheaders
            .EntireColumn.AutoFit
            .Font.color = 16777215
            .Font.Bold = True
            .Interior.color = 5855577
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End With
        'FORMAT THIS AS TEXT, BUT NOT SURE WHY
        With .Range("C:J")
            .NumberFormat = "@"
        End With
    End With
    
    
    itemrow = 4
    'trenddatarow = 2
    
    Sheets("MasterQTO_flat").Activate
    
    'SORT THE FLAT QTO BY AREA>FACILITY>ASSEMBLY CODE>ITEM
    With Sheets("MasterQTO_flat")
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=Cells(3, areacol)
        .Sort.SortFields.Add Key:=Cells(3, facilitycol)
        .Sort.SortFields.Add Key:=Cells(3, assemblycol)
        .Sort.SortFields.Add Key:=Cells(3, itemcol)
        .Sort.SetRange Range("A3:" & lastitem)
        .Sort.Header = xlYes
        .Sort.Apply
    End With
     
'I COULDN'T THINK OF THE BEST WAY TO DO THIS RIGHT NOW, SO I SETTLED WITH A NUMBER OF NESTED IF STATEMENTS TO VERIFY THAT THE ITEM BELONGS AND TO CREATE THE GROUPS ALONG THE WAY
    With Sheets("MasterQTO_flat")
        'RUN THROUGH EACH AREA ROW
        For Each area In Range(Cells(4, areacol), Cells(lastrow, areacol))
            If area <> pArea Then
                'TOGGLE NOT ASSIGNED REMOVAL - CURRENTLY ON
                If area <> "Not Assigned" Then
                    With Sheets("MasterQTO")
                        .Cells(itemrow, 1).Value = "Area " & area
                        .Cells(itemrow, 1).Font.Bold = True
                        .Range("A" & itemrow & ":" & Cells(itemrow, lastcol).Address).Interior.color = 8421504
                    End With
    
                    pArea = area
                    pAreaRow = itemrow
                    itemrow = itemrow + 1
                    
                    'RUN THROUGH EACH FACILITY ROW
                    For Each facility In Range(Cells(4, facilitycol), Cells(lastrow, facilitycol))
                        If facility <> pFacility Then
                            'TOGGLE NOT ASSIGNED REMOVAL - CURRENTLY ON
                            If facility <> "Not Assigned" Then
                                If Cells(facility.Row, areacol).Value = area Then
                                    With Sheets("MasterQTO")
                                        .Cells(itemrow, 1).Value = "Facility " & facility
                                        .Cells(itemrow, 1).Font.Bold = True
                                        .Cells(itemrow, 1).IndentLevel = 1
                                        .Range("A" & itemrow & ":" & Cells(itemrow, lastcol).Address).Interior.color = 10921638
                                    End With
                                    pFacility = facility
                                    pFacilityRow = itemrow
                                    itemrow = itemrow + 1
                                    
                                    'RUN THROUGH EACH ASSEMBLY CODE ROW
                                    For Each assemblycode In Range(Cells(4, assemblycol), Cells(lastrow, assemblycol))
                                        If assemblycode <> pAssemblyCode Then
                                            'TOGGLE NOT ASSIGNED REMOVAL - CURRENTLY ON
                                            If assemblycode <> "Not Assigned" Then
                                                If Cells(assemblycode.Row, areacol).Value = area Then
                                                    If Cells(assemblycode.Row, facilitycol).Value = facility Then
                                                        With Sheets("MasterQTO")
                                                            .Cells(itemrow, 1).Value = assemblycode
                                                            .Cells(itemrow, 1).Font.Bold = True
                                                            .Cells(itemrow, 1).IndentLevel = 2
                                                            
                                                            'REMOVE THIS ROW
                                                            .Cells(itemrow, 26).Value = "Area " & area & ", Facility " & facility & ", " & assemblycode
                                                            
                                                            .Range("a" & itemrow & ":" & Cells(itemrow, lastcol).Address).Interior.color = 14277081
                                                        End With
                                                        pAssemblyCode = assemblycode
                                                        pAssemblyRow = itemrow
                                                        itemrow = itemrow + 1
                                                        
                                                        'RUN THROUGH EACH ITEM ROW
                                                        For Each item In Range(Cells(4, itemcol), Cells(lastrow, itemcol))
                                                            itemlong = assemblycode & "-" & item
                                                            If itemlong <> pitemlong Then
                                                                If Cells(item.Row, areacol).Value = area Then
                                                                    If Cells(item.Row, facilitycol).Value = facility Then
                                                                        If Cells(item.Row, assemblycol).Value = assemblycode Then
                                                                            With Sheets("MasterQTO")
                                                                                .Cells(itemrow, 1).Value = item
                                                                                .Cells(itemrow, 1).IndentLevel = 3
                                                                                For y = 2 To UBound(columnheaders)
                                                                                    itemvalue = columnheaders(y)
                                                                                    If itemvalue <> "Item" And itemvalue <> "Responsible Function" Then
                                                                                        itemvaluecol = Application.Match(columnheaders(y), Worksheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Columns(2), False)
                                                                                        If itemvalue = "Quantity" Then
                                                                                            'IF THERE ARE QUOTES IN THE ASSEMBLY CODE OR DESCRIPTION, REPLACE THEM WITH DOUBLE QUOTES TO BE PLACED IN FORMULAS
                                                                                            assemblycodeformula = Replace(assemblycode, Chr(34), Chr(34) & Chr(34))
                                                                                            itemformula = Replace(item, Chr(34), Chr(34) & Chr(34))
                                                                                            
                                                                                            itemvaluecol = Application.Match(columnheaders(y), Worksheets("CONFIG").ListObjects("QTO_CONFIG").DataBodyRange.Columns(2), False)
                                                                                            
                                                                                            .Cells(itemrow, y).Value = "=SumIfs(MasterQTO_flat!" & .Cells(4, itemvaluecol).Address & ":" & .Cells(lastrow, itemvaluecol).Address & ", MasterQTO_flat!" & .Cells(4, areacol).Address & ":" & .Cells(lastrow, areacol).Address & ", """ & area & """, MasterQTO_flat!" & .Cells(4, facilitycol).Address & ":" & .Cells(lastrow, facilitycol).Address & ", """ & facility & """, MasterQTO_flat!" & .Cells(4, assemblycol).Address & ":" & .Cells(lastrow, assemblycol).Address & ", """ & assemblycodeformula & """,MasterQTO_flat!" & .Cells(4, itemcol).Address & ":" & .Cells(lastrow, itemcol).Address & ", """ & itemformula & """)"
                                                                                            quantitycol = y
                                                                                        Else
                                                                                            .Cells(itemrow, y).Value = Sheets("MasterQTO_flat").Cells(item.Row, itemvaluecol).Value
                                                                                            .Cells(pAssemblyRow, y).Value = Sheets("MasterQTO_flat").Cells(item.Row, itemvaluecol).Value
                                                                                        End If
                                                                                    End If
                                                                                    If itemvalue = "Responsible Function" Then
                                                                                        .Cells(pAssemblyRow, y).Value = "=VLOOKUP(D" & pAssemblyRow & ",'Assembly Codes & Unit Costs'!A:H,8,FALSE)"
                                                                                    End If
                                                                                Next
                                                                            End With
                                                                            pItem = item
                                                                            pitemlong = itemlong
                                                                            pItemRow = itemrow
                                                                            itemrow = itemrow + 1
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                        If pAssemblyRow <> itemrow Then
                                                            'INSERT SUM FOR COST CODE AND GROUP THE ITEM ROWS FOR EACH COST CODE
                                                            With Sheets("MasterQTO")
                                                                sumstart = .Cells(pAssemblyRow + 1, quantitycol).Address
                                                                sumend = .Cells(itemrow - 1, quantitycol).Address
                                                                .Cells(pAssemblyRow, quantitycol).Value = "=SUBTOTAL(9," & sumstart & ":" & sumend & ")"
                                                                .Cells(pAssemblyRow, quantitycol).Font.Bold = True
                                                                
                                                                
                                                                
                                                                'DROP THE VALUES IN THE TREND DATA SHEET
'                                                                Sheets("TrendData").Cells(trenddatarow, 1).Value = "Area " & area & ", Facility " & facility & ", " & assemblycode
'
'                                                                Sheets("TrendData").Cells(trenddatarow, 1).ClearComments
'                                                                Sheets("TrendData").Cells(trenddatarow, 1).AddComment
'                                                                Sheets("TrendData").Cells(trenddatarow, 1).Comment.Text Text:=uom
'
'                                                                'Sheets("TrendData").Cells(trenddatarow, 2).Value = facility
'                                                                'Sheets("TrendData").Cells(trenddatarow, 3).Value = assemblycode
'                                                                trenddatarow = trenddatarow + 1
                                                                
                                                                .Range("A" & pAssemblyRow + 1 & ":" & "A" & itemrow - 1).EntireRow.Group
                                                                .Outline.SummaryRow = False
                                                            End With
                                                        End If
                                                    End If
                                                End If
                                            'TOGGLE NOT ASSIGNED REMOVAL - CURRENTLY ON
                                            End If
                                        End If
                                    Next
                                    If pFacilityRow <> itemrow Then
                                        With Sheets("MasterQTO")
                                            .Range("A" & pFacilityRow + 1 & ":" & "A" & itemrow - 1).EntireRow.Group
                                            .Outline.SummaryRow = False
                                        End With
                                    End If
                                End If
                            'TOGGLE NOT ASSIGNED REMOVAL - CURRENTLY ON
                            End If
                        End If
                    Next
                    If pAreaRow <> itemrow Then
                        With Sheets("MasterQTO")
                            .Range("A" & pAreaRow + 1 & ":" & "A" & itemrow - 1).EntireRow.Group
                            .Outline.SummaryRow = False
                        End With
                    End If
                'TOGGLE NOT ASSIGNED REMOVAL - CURRENTLY ON
                End If
            End If
        Next
    End With
    
    With Sheets("MasterQTO")
        .Columns(1).EntireColumn.AutoFit
        .Outline.ShowLevels rowlevels:=3
    End With
    
Sheets("MasterQTO").Activate
Application.ScreenUpdating = True

End Sub

Sub pmreport()
Application.ScreenUpdating = False
    Dim tol As Variant
    Dim lastrow As Long
    Dim isdimensioned As Boolean
    Dim deleterows() As String
    Dim tolcol As Variant
    Dim toladdr As String
    Dim item As Variant
    Dim indentis As Integer
    Dim thisgroup As String
    Dim prevgroup As String
    Dim prevrow As Long
    Dim x As Variant
    Dim loopcount As Integer
    Dim deleterowsfound As Boolean
    Dim y As Variant
    Dim masterQTOFile As Workbook
    Dim costcode As String
    Dim costcodelong() As String
    Dim usecostcode As Variant
    Dim blockstart As Variant
    Dim blockend As Variant
    
    Set masterQTOFile = ActiveWorkbook
    
    tolcol = Application.Match("Within Tolerance", Sheets("CONFIG").ListObjects("COMPARE_CONFIG").ListColumns("Column Name").DataBodyRange, False)
    Sheets("Comparison").Copy
    ActiveSheet.Name = "PM Summary"
    With Sheets("PM Summary")
        'NEED TO REMOVE STATIC COLUMN REFERENCE
        'RUN THROUGH THE LIST THREE TIMES TO MAKE SURE ALL GROUPS ARE REMOVED. NOT IDEAL AND SHOULD BE CHANGED BUT WORKS TO REMOVE AREA FACILITY AND CODES
        lastrow = .Cells(Rows.count, 10).End(xlUp).Row
        blockend = .Cells(lastrow, 1).Row
            
            'REMOVE THE COST CODE GROUPS THAT ARE NOT SELECTED FOR PM REPORT
            For y = lastrow To 4 Step -1 'Each tol In .Range(Cells(4, tolcol), Cells(lastrow, tolcol))
                If .Cells(y, 1).IndentLevel = 0 Then
                    costcode = .Cells(y, 1).Value
                    costcode = Replace(costcode, "Cost Code ", "")
                    usecostcode = Application.Match(costcode, masterQTOFile.Sheets("CONFIG").ListObjects("PMCODE_CONFIG").ListColumns("Cost Code").DataBodyRange, False)
                        If IsError(usecostcode) Then
                            blockstart = .Cells(y, 1).Row
                            .Rows(blockstart & ":" & blockend).EntireRow.Delete
                            blockend = .Cells(y, 1).Row - 1
                        Else
                            blockend = .Cells(y, 1).Row - 1
                        End If
                End If
            Next
            
            'REMOVE THE LINES THAT ARE WITHIN TOLERENCE
            lastrow = .Cells(Rows.count, 10).End(xlUp).Row
            For y = lastrow To 4 Step -1 'Each tol In .Range(Cells(4, tolcol), Cells(lastrow, tolcol))
                If .Cells(y, 1).IndentLevel = 3 Then
                    tol = .Cells(y, tolcol).Value
                    If tol = "Yes" Then
                        .Cells(y, tolcol).EntireRow.Delete
                    End If
                End If
            Next
            
            
            
'            For y = lastrow To 4 Step -1 'Each tol In .Range(Cells(4, tolcol), Cells(lastrow, tolcol))
'                If .Cells(y, 1).IndentLevel = 3 Then
'                    costcodelong = Split(.Cells(y, 3).Value, Chr(34))
'                    costcode = costcodelong(1)
'                    tol = .Cells(y, tolcol).Value
'                    If tol = "Yes" Then
'                        .Cells(y, tolcol).EntireRow.Delete
'                    End If
'                End If
'            Next
        
        'FIND GROUPING ROWS THAT NEED TO BE REMOVED AND THEN DELETE ENTIRE ROW
'        For loopcount = 1 To 3
'        deleterowsfound = False
'        isdimensioned = False
'        lastrow = .Cells(Rows.Count, 1).End(xlUp).Row
'            For Each item In .Range("A4:A" & lastrow)
'            indentis = item.IndentLevel
'                If indentis = 0 Then thisgroup = "CostCode"
'                If indentis = 1 Then thisgroup = "Area"
'                'If indentis = 2 Then thisgroup = "Lvl3Code"
'                If indentis = 3 Then thisgroup = "Facility"
'                If thisgroup <> "Facility" Then
'                    If thisgroup <= prevgroup Then
'                        If prevgroup <> "Facility" Then
'                            If isdimensioned = True Then
'                                ReDim Preserve deleterows(1 To UBound(deleterows) + 1) As String
'                            Else
'                                ReDim deleterows(1 To 1) As String
'                                isdimensioned = True
'                            End If
'                            deleterows(UBound(deleterows)) = prevrow
'                            deleterowsfound = True
'                        End If
'                    End If
'                End If
'                prevgroup = thisgroup
'                prevrow = item.Row
'            Next
'            'Add the last row on
'            If prevgroup <> "Facility" Then
'                If isdimensioned = True Then
'                    ReDim Preserve deleterows(1 To UBound(deleterows) + 1) As String
'                Else
'                    ReDim deleterows(1 To 1) As String
'                    isdimensioned = True
'                End If
'                deleterows(UBound(deleterows)) = prevrow
'                deleterowsfound = True
'            End If
'            If deleterowsfound = True Then
'                For x = UBound(deleterows) To LBound(deleterows) Step -1
'                    '.Rows(deleterows(x)).EntireRow.Interior.Color = 5046
'                    .Rows(deleterows(x)).EntireRow.Delete
'                Next
'            End If
'            Erase deleterows
'        Next
    End With
    
Application.ScreenUpdating = True
End Sub








