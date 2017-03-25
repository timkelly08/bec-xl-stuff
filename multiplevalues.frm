VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} multiplevalues 
   Caption         =   "Select Search Criteria"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7755
   OleObjectBlob   =   "multiplevalues.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "multiplevalues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.width) - (0.5 * Me.width)
    Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
End Sub
Private Sub cancelbutton_Click()
    Me.Hide
End Sub

Private Sub addcriteriabutton_Click()
    For i = AvailableValuesListBox.ListCount - 1 To 0 Step -1
        If AvailableValuesListBox.Selected(i) = True Then
            searchval = AvailableValuesListBox.List(i)
            AvailableValuesListBox.RemoveItem i
            SelectedValuesListBox.AddItem searchval
        End If
    Next i
    
    With SelectedValuesListBox
        For j = 0 To SelectedValuesListBox.ListCount - 2
            For i = 0 To SelectedValuesListBox.ListCount - 2
                If .List(i) > .List(i + 1) Then
                    temp = .List(i)
                    .List(i) = .List(i + 1)
                    .List(i + 1) = temp
                End If
            Next i
        Next j
    End With
    
End Sub
Private Sub removecriteriabutton_Click()
    For i = SelectedValuesListBox.ListCount - 1 To 0 Step -1
        If SelectedValuesListBox.Selected(i) = True Then
            searchval = SelectedValuesListBox.List(i)
            SelectedValuesListBox.RemoveItem i
            AvailableValuesListBox.AddItem searchval
        End If
    Next i
    
    With AvailableValuesListBox
        For j = 0 To AvailableValuesListBox.ListCount - 2
            For i = 0 To AvailableValuesListBox.ListCount - 2
                If .List(i) > .List(i + 1) Then
                    temp = .List(i)
                    .List(i) = .List(i + 1)
                    .List(i + 1) = temp
                End If
            Next i
        Next j
    End With
End Sub

Private Sub runcollapse_Click()
collapseform.Hide

itemcolumn = 6
groupbycount = itemcolumn - 2
AssembleExport = "Untitled View"
headerrow = 4
lastrow = 1486

If Not groupbycount = 0 Then
    For x = 2 To groupbycount + 1
        columnname1 = Worksheets(AssembleExport).Cells(headerrow, x).Value
        firstgb = Worksheets(AssembleExport).Cells(headerrow + 1, x).Address
        lastgb = Worksheets(AssembleExport).Cells(lastrow, x).Address
            For Each gb In Sheets(AssembleExport).Range(firstgb & ":" & lastgb)
                gbaddr = gb.Address
                gbmain = ("F" & gb.Row)
                If Sheets(AssembleExport).Range(gbaddr).Font.color = 0 Then
                    If gb <> "" Then
                        Sheets(AssembleExport).Range(gbmain).Value = gb
                        Sheets(AssembleExport).Range(gbmain).Font.color = 0
                    End If
                End If
            Next
    Next x
    Sheets(AssembleExport).Range("A1:E1").EntireColumn.Delete
    Sheets(AssembleExport).Activate
End If
End Sub

Private Sub SaveButton_Click()
    If Me.SelectedValuesListBox.ListIndex <> -1 Then
        createformula.ValueList1.List = Me.SelectedValuesListBox.List
        
        For i = createformula.ValueCombo1.ListCount - 1 To 0 Step -1
            createformula.ValueCombo1.RemoveItem i
        Next
        createformula.ValueCombo1.AddItem "<Multiple Values Selected>"
        createformula.ValueCombo1.AddItem "<Edit Selection>"
        createformula.ValueCombo1.AddItem "<Clear Selection>"
        createformula.ValueCombo1.ListIndex = 0
        For i = Me.SelectedValuesListBox.ListCount - 1 To 0 Step -1
            Me.SelectedValuesListBox.RemoveItem i
        Next
        multiplevalues.Hide

    End If
End Sub



