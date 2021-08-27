Attribute VB_Name = "InterventionListBoxes"
Sub PopulateListBox(InterventionList As Worksheet, LS As MSForms.ListBox)
'ataSheet As Worksheet
Dim LastRow As Long
Dim objCell As Range
    
    LS.Clear
    
    LastRow = InterventionList.Cells(Rows.Count, "B").End(xlUp).Row
    For Each objCell In InterventionList.Range("B2:B" & LastRow)
       If objCell <> "" Then
        LS.AddItem objCell.Value
       End If
Next objCell

End Sub
Sub PopList(InterventionList As Worksheet, LS As MSForms.ListBox)
'ataSheet As Worksheet
Dim LastRow As Long
Dim objCell As Range
    
With LS
    .ColumnCount = 3
    .ColumnWidths = "50;100;0"
End With

LastRow = InterventionList.Cells(Rows.Count, "B").End(xlUp).Row
For i = 0 To LastRow
j = i + 2
    LS.AddItem
    LS.List(i, 0) = InterventionList.Cells(j, 1)
    LS.List(i, 1) = InterventionList.Cells(j, 2)
    LS.List(i, 2) = InterventionList.Cells(j, 4)

Next i

End Sub
Sub PopuList(LSAll As MSForms.ListBox, LSCur As MSForms.ListBox)
    Dim i As Integer

    For i = 0 To LSAll.ListCount - 1
        If LSAll.Selected(i) = True Then
            LSCur.AddItem LSAll.List(i)

        End If
    Next i

End Sub
Sub AddCurrent(LSAll As MSForms.ListBox, LSCur As MSForms.ListBox)
    Dim i As Integer

    For i = 0 To LSAll.ListCount - 1
        If LSAll.Selected(i) = True Then
            LSCur.AddItem LSAll.List(i)

        End If
    Next i

End Sub
Sub AddOpp(LSAll As MSForms.ListBox, LSOpp As MSForms.ListBox)

    Dim i As Integer

    For i = 0 To LSAll.ListCount - 1
        If LSAll.Selected(i) = True Then
            LSOpp.AddItem LSAll.List(i)
        End If
    Next i

End Sub

Sub RemoveCur(LSCur As MSForms.ListBox)

    Dim i As Integer

For i = 0 To LSCur.ListCount - 1
    If LSCur.Selected(i) Then
        LSCur.RemoveItem (i)
    End If
Next i

End Sub
Sub RemoveOpp(LSOpp As MSForms.ListBox)

    Dim i As Integer

For i = 0 To LSOpp.ListCount - 1
    If LSOpp.Selected(i) Then
        LSOpp.RemoveItem (i)
    End If
Next i

End Sub





