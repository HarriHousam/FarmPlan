Attribute VB_Name = "FarmPlan"

Sub PrintFarmPlan(Farm As Worksheet)


'Declare Variables

   Dim WordApp As Word.Application
   Dim wDoc As Word.Document
   Dim ws As Worksheet
   Dim ObjShape As InlineShape
   Dim recList
    Dim rec As Variant
    Dim i As Integer
   Dim objCC As ContentControl
    Dim dummyText As String
    Dim Bookmark1 As Range
    Dim pctdone As Single
    Dim cellRange1 As String
    Dim cellRange As Range
    
'Display Progress Bar
LoadingBar.Bar.Width = 0
LoadingBar.Show

    
    
'Setup working area
Set WordApp = New Word.Application

    Set wDoc = WordApp.Documents.Add(Template:="C:\Users\HarrietHousam\OneDrive - Westcountry Rivers Trust\Documents\UST3\Farm Plan Think Peice\Farm Plan WRT logo.dotm", NewTemplate:=False, DocumentType:=0)
    WordApp.Visible = True
    
    Set ws = Farm
    
FractionComplete (0)
'Complete Title Page and Common Tags

    For Each objCC In wDoc.SelectContentControlsByTag("ProjectName")
        objCC.Range.Text = ws.Range("E4").Value
    Next
    For Each objCC In wDoc.SelectContentControlsByTag("FarmName")
        objCC.Range.Text = ws.Range("E20").Value
    Next
    
    
    For Each objCC In wDoc.SelectContentControlsByTag("AdvisorName")
        objCC.Range.Text = ws.Range("I4").Value
    Next
    
        For Each objCC In wDoc.SelectContentControlsByTag("FarmNumber")
        objCC.Range.Text = ws.Range("E20").Value
    Next
    
    
    For Each objCC In wDoc.SelectContentControlsByTag("FarmSize")
        objCC.Range.Text = ws.Range("E36").Value
    Next

FractionComplete (0.25)
'Create list of recommendations
     
    Set recList = FarmRecomms(Farm)

'Go to Recommendations Bookmark
'    WordApp.ActiveDocument.Selection.Goto What:=wdGoToBookmark, Name:="Recommendations"
Set Bookmark = WordApp.ActiveDocument.Bookmarks("Recommendations").Range
'Loop through recommendaions to create headers
i = 1
Set Bookmark = WordApp.ActiveDocument.Bookmarks("Recommendations").Range
 '   With WordApp.ActiveDocument.Bookmarks("Recommendations").Range
        For Each rec In recList.Keys()
            Set cellRange = Range(recList(rec))
    Selection.Font.Bold = wdToggle
    Selection.Font.Size = 12
    Selection.TypeText Text:="Recommendation " & i & ": " & rec & vbNewLine
    Selection.Font.Bold = wdToggle
    Selection.Font.Size = 11
    Selection.TypeText Text:=cellRange.Offset(, 1).Value & vbNewLine

    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=3, NumColumns:= _
        3, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    With Selection.Tables(1)
        If .Style <> "Table Grid" Then
            .Style = "Table Grid"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With

        .InsertAfter Text:="Recommendation " & i & ": " & rec & vbNewLine
        With Selection.Font
            .Name = "Calibre"
            .Size = 11
            .Bold = False
        End With
        
                .InsertAfter Text:=cellRange.Offset(, 1).Value & vbNewLine

        .Tables.Add Range:=Bookmark, NumRows:=2, NumColumns:=2
        
        
        
'        "Benefit:" & cellRange.Offset(, 2).Value & vbNewLine
        i = i + 1


      '  .Bookmarks("Recommendations").Range.Text = "Recommendation " & i & ": " & rec & vbNewLine

         'insertparagraphafer.cellRange.Offset(, 1).Value & vbNewLine & "Benefit:" & cellRange.Offset(, 2).Value & vbNewLine

    Next
 
    End With

FractionComplete (0.5)

'Complete and Clean Up

Set wDoc = Nothing

FractionComplete (1)

Unload LoadingBar


End Sub

Sub PrintThisFarmPlan()

Dim ws As Worksheet
Set ws = ActiveSheet
Call PrintFarmPlan(ws)
End Sub


Sub FractionComplete(pctdone As Single)
With LoadingBar
    .LoadingCaption.Caption = pctdone * 100 & "% Complete"
    .Bar.Width = pctdone * (.LoadingFrame.Width)
End With
DoEvents
End Sub
