Attribute VB_Name = "FunctionalityButtons"
Option Explicit

Sub NewFarmForm()

' NewFarmForm Macro
' Keyboard Shortcut: Ctrl+n
' Copies the blank template (hidden sheeet) and renames the new sheet as 'Farm' plus the next farm number

    Dim NewFarmValue As Double
    Dim SheetName As String
    
    Sheets("Farm Checklist Original").Visible = True
    Sheets("Farm Checklist Original").Copy Before:=Sheets("Farm Checklist Original")
    NewFarmValue = Sheets("Report Builder").Range("B9").Value
    NewFarmValue = NewFarmValue + 1
    SheetName = "Farm " & NewFarmValue
    ActiveSheet.Name = SheetName
    Sheets("Report Builder").Range("B9").Value = NewFarmValue
    
    Worksheets("Report Builder").Select
    Range("B9").Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, address:="", SubAddress:= _
        "'" & SheetName & "'!R1C1"
        
    Sheets("Farm Checklist Original").Visible = False
End Sub

Sub ProduceOverview()


'Loops through sheets whose name beings with "Farm"
'Adds Ha

Dim ws As Worksheet
Dim HaSum As Integer
Dim FarmCounter As Integer
Dim RecCol As Collection
Dim Recom As Variant
Dim objCell As Range
Dim WQCellRange As Integer



    HaSum = 0
    FarmCounter = 0
    WQCellRange = 13
    For Each ws In Sheets 'This statement starts the loop
        If ws.Name Like "Farm*" Then 'Perform the Excel action you wish
            If Not IsEmpty(ws.Range("E20").Value) Then
                HaSum = HaSum + ws.Range("E36").Value
                FarmCounter = FarmCounter + 1
                'Compare Staked Info lists to Farm Recommendations to look for opportunities
                Set RecCol = FarmRecomms(ws)
                For Each Recom In RecCol
                    For Each objCell In Sheets("StackedInfo").Range("C4:C10")
                        If objCell = Recom Then
                            Sheets("StackedInfo").Range("C" & WQCellRange).Value = ws.Name
                            Exit For
                            Exit For
                            WQCellRange = WQCellRange + 1
                        End If
                    Next objCell
                Next Recom
            End If
        End If
    Next ws

    

Worksheets("Report Builder").Select
Range("F6").Value = FarmCounter
Range("H6").Value = HaSum



End Sub

Sub ProduceTableOverview()

Dim startRow, startCol, LastRow, lastCol As Long
Dim headers As Range

'Set Master sheet for consolidation
Set mtr = Worksheets("Master")

Set wb = ThisWorkbook
'Get Headers
Set headers = Application.InputBox("Select the Headers", Type:=8)

'Copy Headers into master
headers.Copy mtr.Range("A1")
startRow = headers.Row + 1
startCol = headers.Column

Debug.Print startRow, startCol
'loop through all sheets
For Each ws In wb.Worksheets
     'except the master sheet from looping
     If ws.Name <> "Master" Then
        ws.Activate
        LastRow = Cells(Rows.Count, startCol).End(xlUp).Row
        lastCol = Cells(startRow, Columns.Count).End(xlToLeft).Column
        'get data from each worksheet and copy it into Master sheet
        Range(Cells(startRow, startCol), Cells(LastRow, lastCol)).Copy _
        mtr.Range("A" & mtr.Cells(Rows.Count, 1).End(xlUp).Row + 1)
           End If
Next ws

Worksheets("Master").Activate

End Sub

End Sub


Sub OpportunityMap()


'Loops through sheets whose name beings with "Farm"

Dim ws As Worksheet
Dim HaSumWill As Integer
Dim FarmsOpp As Integer
Dim FarmsWill As Integer



    HaSumWill = 0
    FarmsOpp = 0
    FarmsWill = 0

    For Each ws In Sheets 'This statement starts the loop
        If ws.Name Like "Farm*" Then 'Perform the Excel action you wish
            If ws.Range("c18") = "opportunity but not willing" Or ws.Range("c18") = "Opportunity and Willing" Then
            FarmsOpp = FarmsOpp + 1
            End If
            If ws.Range("c18") = "Opportunity and Willing" Then
            FarmsWill = FarmsWill + 1
            HaSumWill = HaSumWill + ws.Range("C8").Value
            End If
        End If
    Next ws


Worksheets("Report Builder").Select
Range("F12").Value = FarmsOpp
Range("G12").Value = FarmsWill
Range("H12").Value = HaSumWill




End Sub

