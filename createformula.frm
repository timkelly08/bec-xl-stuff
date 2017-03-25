VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} createformula 
   Caption         =   "Create New Rule"
   ClientHeight    =   6825
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



Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.width) - (0.5 * Me.width)
    Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
    columnheaders = Sheets("CONFIG").ListObjects("QTO_CONFIG").ListColumns(2).DataBodyRange
    PropertyCombo.List = columnheaders
    
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

Private Sub CreateFormulaButton_Click()
newrow = ThisWorkbook.Sheets("RULES").Cells(Rows.count, "A").End(xlUp).Row + 1
newname = FormulaName.Text
newpropertyname = PropertyCombo.Text
newpropertyvalue = PropertyValue.Text
newformula = FormulaBox.Text
newcostcode = costcode.Text
newuom = uom.Text
replaceQ = replaceQTY.Value

ThisWorkbook.Sheets("RULES").Range("A" & newrow).Value = newname
ThisWorkbook.Sheets("RULES").Range("B" & newrow).Value = newcostcode
ThisWorkbook.Sheets("RULES").Range("C" & newrow).Value = newpropertyname
ThisWorkbook.Sheets("RULES").Range("D" & newrow).Value = newpropertyvalue
ThisWorkbook.Sheets("RULES").Range("E" & newrow).Value = newuom
ThisWorkbook.Sheets("RULES").Range("F" & newrow).Value = newformula
ThisWorkbook.Sheets("RULES").Range("G" & newrow).Value = replaceQ

createformula.Hide
End
End Sub


Private Sub CancelButton_Click()
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
