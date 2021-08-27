Attribute VB_Name = "ListBoxToText"

Function FarmRecomms(Farm As Worksheet) As Object

Dim Size As Integer
Dim dict
Dim i As Integer
Dim listValue As String
Dim LastRow As Integer
Dim LB As OLEObject
Dim WaterLB As OLEObject
Dim sheet As Worksheet
Dim address As Range
Dim cellRef As String

Set dict = CreateObject("Scripting.Dictionary")

'Loop through each list box and add to a 'master' list containing all the current recommendations

'Metric 1
Set sheet = Sheets("InterventionsInfrastructure")
Set LB = Farm.OLEObjects("InfCur")
Size = LB.Object.ListCount - 1
'loop through current recommendations listbox. Add text value and cell reference to a dictionary
    For i = 0 To Size
        LastRow = sheet.Cells(Rows.Count, "B").End(xlUp).Row
        Set address = sheet.Range("B2:B" & LastRow).Find(What:=LB.Object.List(i), LookIn:=xlValues)
        cellRef = "" & sheet.Name & "!" & address.address
        dict.Add LB.Object.List(i), cellRef
    Next i
    
MsgBox dict.Count

'metric 2
Set sheet = Sheets("InterventionsWater")
Set LB = Farm.OLEObjects("WaterCurrent")
Size = LB.Object.ListCount - 1
    For i = 0 To Size
        LastRow = sheet.Cells(Rows.Count, "B").End(xlUp).Row
        Set address = sheet.Range("B2:B" & LastRow).Find(What:=LB.Object.List(i), LookIn:=xlValues)
        cellRef = "" & sheet.Name & "!" & address.address
        dict.Add LB.Object.List(i), cellRef
    Next i

Set FarmRecomms = dict

End Function

Sub InsertRecommsText()
 Dim rngFormat As Range
 Set rngFormat = ActiveDocument.Range(Start:=0, End:=0)
 With rngFormat
 .InsertAfter Text:="Title"
 .InsertParagraphAfter
 With .Font
 .Name = "Tahoma"
 .Size = 24
 .Bold = True
 End With
 End With
 With ActiveDocument.Paragraphs(1)
 .Alignment = wdAlignParagraphCenter
 .SpaceAfter = InchesToPoints(0.5)
 End With
End Sub
