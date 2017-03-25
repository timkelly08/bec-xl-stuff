VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} createformula 
   Caption         =   "Create New Rule"
   ClientHeight    =   11070
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   9240.001
   OleObjectBlob   =   "createformula.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "createformula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim y As Integer
Dim currentfield As String

Private Sub Image2_Click()
ValueCombo1.Clear
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.width) - (0.5 * Me.width)
    Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
    columnheaders = Sheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns(2).DataBodyRange
    
        PropertyCombo1.List = columnheaders
        PropertyCombo2.List = columnheaders
        PropertyCombo3.List = columnheaders
        PropertyCombo4.List = columnheaders
        PropertyCombo5.List = columnheaders
        
    'SET UP THE START POSITION FOR FORM RELATED TO THE SEARCH CRITERIA GROUP
        searchcriteria.height = 65
        shiftgroup.Top = 170
        Me.height = 425
        
        y = 1
        
    
    'SET UP THE FORMULA QUANTITIES
    x = 0
    For Each isQTY In Worksheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns("IsQuantity?").DataBodyRange
        If isQTY = True Then
            QTY = Worksheets("CONFIG").Cells(isQTY.Row, 2).Value
            x = x + 1
            If x = 1 Then
                createformula.qty1.Visible = True
                createformula.qty1.Caption = QTY
            End If
            If x = 2 Then
                createformula.qty2.Visible = True
                createformula.qty2.Caption = QTY
            End If
            If x = 3 Then
                createformula.qty3.Visible = True
                createformula.qty3.Caption = QTY
            End If
            If x = 4 Then
                createformula.qty4.Visible = True
                createformula.qty4.Caption = QTY
            End If
            If x = 5 Then
                createformula.qty5.Visible = True
                createformula.qty5.Caption = QTY
            End If
            If x = 6 Then
                createformula.qty6.Visible = True
                createformula.qty6.Caption = QTY
            End If
            If x = 7 Then
                createformula.qty7.Visible = True
                createformula.qty7.Caption = QTY
            End If
            If x = 8 Then
                createformula.qty8.Visible = True
                createformula.qty8.Caption = QTY
            End If
            If x = 9 Then
                createformula.qty9.Visible = True
                createformula.qty9.Caption = QTY
            End If
            If x = 10 Then
                createformula.qty10.Visible = True
                createformula.qty10.Caption = QTY
            End If
            If x = 11 Then
                createformula.qty11.Visible = True
                createformula.qty11.Caption = QTY
            End If
            If x = 12 Then
                createformula.qty12.Visible = True
                createformula.qty12.Caption = QTY
            End If
        End If
    Next

End Sub

Private Sub AddSearchCriteria_Click()
    y = y + 1
    
    If y = 2 Then
        searchcriteria.height = searchcriteria.height + 25
        shiftgroup.Top = shiftgroup.Top + 25
        Me.height = Me.height + 25
        PropertyName2.Visible = True
        PropertyCombo2.Visible = True
        ValueLabel2.Visible = True
        ValueCombo2.Visible = True
    End If
    If y = 3 Then
        searchcriteria.height = searchcriteria.height + 25
        shiftgroup.Top = shiftgroup.Top + 25
        Me.height = Me.height + 25
        PropertyName3.Visible = True
        PropertyCombo3.Visible = True
        ValueLabel3.Visible = True
        ValueCombo3.Visible = True
    End If
    If y = 4 Then
        searchcriteria.height = searchcriteria.height + 25
        shiftgroup.Top = shiftgroup.Top + 25
        Me.height = Me.height + 25
        PropertyName4.Visible = True
        PropertyCombo4.Visible = True
        ValueLabel4.Visible = True
        ValueCombo4.Visible = True
    End If
    If y = 5 Then
        searchcriteria.height = searchcriteria.height + 25
        shiftgroup.Top = shiftgroup.Top + 25
        Me.height = Me.height + 25
        PropertyName5.Visible = True
        PropertyCombo5.Visible = True
        ValueLabel5.Visible = True
        ValueCombo5.Visible = True
        AddSearchCriteria.Visible = False
    End If
End Sub

Private Sub CreateFormulaButton_Click()
newrow = ThisWorkbook.Sheets("RULES").Cells(Rows.count, "A").End(xlUp).Row + 1
newname = FormulaName.Text
newformula = FormulaBox.Text
newcostcode = costcode.Text
newuom = uom.Text
ReplaceQ = replaceQTY.Value

''''''''''''''''''''''''''''''''''
If PropertyCombo1.Text <> "" Then
newpropertyname = PropertyCombo1.Text

    If ValueCombo1.Text <> "<Multiple Values Selected>" Then
        newpropertyvalue = "<Value>" & ValueCombo1.Text & "</Value>"
    Else
        For i = 0 To Me.ValueList1.ListCount - 1
            newpropertyvalue = newpropertyvalue & "<Value>" & Me.ValueList1.List(i) & "</Value>"
        Next
    End If
End If

newsearchcriteria = "<Field><ColumnName>" & newpropertyname & "</ColumnName><Values>" & newpropertyvalue & "</Values></Field>"
'''''''''''''''''''''''''''''''

XMLstring = "<Rule><RuleName>" & newname & "</RuleName><SearchCriteria>" & newsearchcriteria & "</SearchCriteria><UOM>" & newuom _
& "</UOM><CostCode>" & newcostcode & "</CostCode><Formula>" & newformula & "</Formula><Replace>" & ReplaceQ & "</Replace></Rule>"

ThisWorkbook.Sheets("RULES").Range("A" & newrow).Value = XMLstring

createformula.Hide
End Sub


Private Sub cancelbutton_Click()
End
End Sub

Private Sub RunUpdateButton_Click()
newrow = ThisWorkbook.Sheets("RULES").Cells(Rows.count, "A").End(xlUp).Row + 1
newname = FormulaName.Text
newpropertyname = GroupCombo.Text
newpropertyvalue = PropertyValue.Text
newformula = FormulaBox.Text

ThisWorkbook.Sheets("RULES").Range("A" & newrow).Value = newname
ThisWorkbook.Sheets("RULES").Range("B" & newrow).Value = newpropertyname
ThisWorkbook.Sheets("RULES").Range("C" & newrow).Value = newpropertyvalue
ThisWorkbook.Sheets("RULES").Range("D" & newrow).Value = newformula

createformula.Hide

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ADD COLUMNS IN BRACKETS TO FORMULA BOX

Private Sub qty1_Click()
    currentformula = FormulaBox.Text
    FormulaBox.Text = currentformula & "[" & qty1.Caption & "]"
    FormulaBox.SetFocus
End Sub
Private Sub qty2_Click()
    currentformula = FormulaBox.Text
    FormulaBox.Text = currentformula & "[" & qty2.Caption & "]"
    FormulaBox.SetFocus
End Sub
Private Sub qty3_Click()
    currentformula = FormulaBox.Text
    FormulaBox.Text = currentformula & "[" & qty3.Caption & "]"
    FormulaBox.SetFocus
End Sub
Private Sub qty4_Click()
    currentformula = FormulaBox.Text
    FormulaBox.Text = currentformula & "[" & qty4.Caption & "]"
    FormulaBox.SetFocus
End Sub
Private Sub qty5_Click()
    currentformula = FormulaBox.Text
    FormulaBox.Text = currentformula & "[" & qty5.Caption & "]"
    FormulaBox.SetFocus
End Sub
Private Sub qty6_Click()
    currentformula = FormulaBox.Text
    FormulaBox.Text = currentformula & "[" & qty6.Caption & "]"
    FormulaBox.SetFocus
End Sub
Private Sub qty7_Click()
    currentformula = FormulaBox.Text
    FormulaBox.Text = currentformula & "[" & qty7.Caption & "]"
    FormulaBox.SetFocus
End Sub
Private Sub qty8_Click()
    currentformula = FormulaBox.Text
    FormulaBox.Text = currentformula & "[" & qty8.Caption & "]"
    FormulaBox.SetFocus
End Sub
Private Sub qty9_Click()
    currentformula = FormulaBox.Text
    FormulaBox.Text = currentformula & "[" & qty9.Caption & "]"
    FormulaBox.SetFocus
End Sub
Private Sub qty10_Click()
    currentformula = FormulaBox.Text
    FormulaBox.Text = currentformula & "[" & qty10.Caption & "]"
    FormulaBox.SetFocus
End Sub
Private Sub qty11_Click()
    currentformula = FormulaBox.Text
    FormulaBox.Text = currentformula & "[" & qty11.Caption & "]"
    FormulaBox.SetFocus
End Sub
Private Sub qty12_Click()
    currentformula = FormulaBox.Text
    FormulaBox.Text = currentformula & "[" & qty12.Caption & "]"
    FormulaBox.SetFocus
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub PropertyCombo1_Change()
Application.ScreenUpdating = False

Sheets("MasterQTO_flat").Activate
    
    With Me.PropertyCombo1
        currentfield = .List(.ListIndex)
    End With
    
    ValueCombo1.Clear
    For i = ValueCombo1.ListCount - 1 To 0 Step -1
        ValueCombo1.RemoveItem i
    Next i
    
    With Sheets("MasterQTO_flat")
        propertycolumn = Application.Match(currentfield, .Range("A3:AAA3"), False)
        lastrow = .Cells(Rows.count, propertycolumn).End(xlUp).Row
        If Not IsError(propertycolumn) Then
            firstnameaddr = .Cells(4, propertycolumn).Address
            lastnameaddr = .Cells(lastrow, propertycolumn).Address
            If lastrow <> 3 Then 'MAKE SURE THE COLUMN HAS VALUES
                .Range("CC1").Value = "<Select Multiple Items>"
                .Range(firstnameaddr & ":" & lastnameaddr).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=.Range("CC2"), Unique:=True
                .Range(.Range("CC2"), .Range("CC" & lastrow).End(xlUp)).Sort key1:=Range("CC2"), order1:=xlAscending
                .Range("CC2:CC" & lastrow).RemoveDuplicates Columns:=(1), Header:=xlNo
                newend = .Range("CC" & lastrow).End(xlUp).Row
                If newend > 1 Then
                    FieldData1 = .Range(.Range("CC1"), .Range("CC" & newend)).Value
                End If
                .Range("CC:CC").EntireColumn.Delete
            End If
        End If
    End With
    
    If Not IsEmpty(FieldData1) Then
        ValueCombo1.List = FieldData1
    Else
        For i = ValueCombo1.ListCount - 1 To 0 Step -1
            ValueCombo1.RemoveItem i
        Next
    End If
    
Application.ScreenUpdating = True
End Sub

Private Sub ValueCombo1_Change()

    If ValueCombo1.ListIndex <> -1 Then

    With Me.ValueCombo1
            If .List(.ListIndex) = "<Select Multiple Items>" Then
            setmultivalue = True
            End If
    End With
    If setmultivalue = True Then
        multiplevalues.AvailableValuesListBox.List = Me.ValueCombo1.List
        multiplevalues.AvailableValuesListBox.RemoveItem (0)
        multiplevalues.multivalueselect.Caption = "Select Multiple Values for " & currentfield
        multiplevalues.Show

    End If
'If Me.ValueCombo1.ListIndex = "<Select Multiple Items>" Then
'    Debug.Print "Go"
'End If

    End If
End Sub
