VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub InfAll_Click()

End Sub

Private Sub Worksheet_Activate()

Dim sh1 As Worksheet
Dim sh2 As Worksheet
Dim sh3 As Worksheet
Dim PosHeight As Integer
Dim InfHeight As Integer
Dim WaterHeight As Integer
Dim SoilHeight As Integer

'Set Listbox positions

PosHeight = Range("B93").Top
InfHeight = Range("B102").Top
WaterHeight = Range("B113").Top
SoilHeight = Range("B124").Top

With OLEObjects("Positives")
    .Left = 45
    .Height = 200
    .Top = PosHeight
    .Width = 500
End With
With OLEObjects("InfAll")
    .Left = 45
    .Height = 200
    .Top = InfHeight
    .Width = 350
End With
With OLEObjects("InfCur")
    .Left = 500
    .Height = 100
    .Top = InfHeight
    .Width = 350
End With
With OLEObjects("InfraOpp")
    .Left = 500
    .Height = 100
    .Top = InfHeight + 120
    .Width = 350
End With
With OLEObjects("WaterAll")
    .Left = 45
    .Height = 200
    .Top = WaterHeight
    .Width = 350
End With
With OLEObjects("WaterCurrent")
    .Left = 500
    .Height = 100
    .Top = WaterHeight
    .Width = 350
End With
With OLEObjects("WaterPoten")
    .Left = 500
    .Height = 100
    .Top = WaterHeight + 120
    .Width = 350
End With
With OLEObjects("SoilAll")
    .Left = 45
    .Height = 200
    .Top = SoilHeight
    .Width = 350
End With
With OLEObjects("SoilCur")
    .Left = 500
    .Height = 100
    .Top = SoilHeight
    .Width = 350
End With
With OLEObjects("SoilPoten")
    .Left = 500
    .Height = 100
    .Top = SoilHeight + 120
    .Width = 350
End With

'Populate Listboxes
Set sh1 = Sheets("Positives")
Set sh2 = Sheets("InterventionsInfrastructure")
Set sh3 = Sheets("InterventionsWater")

Call PopList(sh1, Me.Positives)
Call PopulateListBox(sh2, Me.InfAll)
Call PopulateListBox(sh3, Me.WaterAll)



End Sub


Private Sub AddCurrInf_Click()

Call AddCurrent(Me.InfAll, Me.InfCur)
End Sub

Private Sub AddOppIN_Click()

Call AddOpp(Me.InfAll, Me.InfraOpp)

End Sub

Private Sub RemoveCurrIN_Click()

Call RemoveCur(Me.InfCur)


End Sub

Private Sub RemoveOppIN_Click()

Call RemoveOpp(Me.InfraOpp)


End Sub


Private Sub AddCurWM_Click()

Call AddCurrent(Me.WaterAll, Me.WaterCurrent)

End Sub
Private Sub AddOppWM_Click()

Call AddOpp(Me.WaterAll, Me.WaterPoten)

End Sub

Private Sub RemoveCurWM_Click()

Call RemoveCur(Me.WaterCurrent)


End Sub

Private Sub RemoveOppWM_Click()

Call RemoveOpp(Me.WaterPoten)

End Sub







