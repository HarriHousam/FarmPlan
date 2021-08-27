VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



Private Sub Worksheet_Activate()

Dim sh As Worksheet

'If the operating system is Windows
If Application.OperatingSystem Like "*Window*" Then
'If Mac or other operating system
Else
    MsgBox "Your opperating system is " & _
    Application.OperatingSystem & ". As this is not Windows you may loose some functionality. Please contact Harriet Housam if you have any concerns or issues."
End If


If Application.Version < 16# Then
    MsgBox "Your Microsoft Excel version is " & _
    Application.Version & ". This spreadsheet was developed on a newer version (16.0). Therefore, you may loose some functionality. Please contact Harriet Housam if you have any concerns or issues."
End If


Me.CBFarmList.Clear
For Each sh In ThisWorkbook.Worksheets
    If sh.Name Like "Farm*" Then 'Perform the Excel action you wish
        Me.CBFarmList.AddItem sh.Name
    End If
Next sh

End Sub



Sub PrintFarmPlanFromSelection()

    If CBFarmList.Value <> "" Then
            Call PrintFarmPlan(Worksheets(CBFarmList.Value))
    End If


End Sub


Sub GoToFarm()

If CBFarmList.Value <> "" Then
    Worksheets(CBFarmList.Value).Select
    End If
    
End Sub
