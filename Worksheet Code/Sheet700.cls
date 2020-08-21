VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub cmdExportCMF1_Click()

    If Me.Range("MatchesKnown1").Offset(2, 0).Value = "" Then Exit Sub
    
    'Specify contributors if the other slot on the CODIS sheet is full
    ExportCMF 1, Application.WorksheetFunction.CountA(Me.Range("Dest_CODIS2").Offset(0, 1)) > 0
    
End Sub

Private Sub cmdExportCMF2_Click()

    If Me.Range("MatchesKnown2").Offset(2, 0).Value = "" Then Exit Sub
    
    'Specify contributors if the other slot on the CODIS sheet is full
    'This phrase equates to T/F: Application.WorksheetFunction.CountA(Me.Range("Dest_CODIS1").Offset(0, 1)) > 0
    ExportCMF 2, Application.WorksheetFunction.CountA(Me.Range("Dest_CODIS1").Offset(0, 1)) > 0
    
End Sub


Sub ExportCMF(Profile As Integer, SpecifyContributors As Boolean)

    Dim answer As Variant

    If Me.Range("CODIS_CaseNum").Value = "" Then Exit Sub
    
    If Me.Range("SendtoCODIS" & Profile).Offset(-2, 0).Value = "" Then
        MsgBox "Please select a specimen category for this sample.", vbOKOnly + vbExclamation, "Specimen Category Required"
        Exit Sub
    End If
    
    If Me.Range("CODIS_Comment").Value = "" Then
        MsgBox "Please enter an item description for this sample.", vbOKOnly + vbExclamation, "Item Description Required"
        Exit Sub
    End If
    
    If Me.Range("MatchesKnown" & Profile).Value = "" Then
        answer = MsgBox("Does this sample match a known profile?", vbYesNoCancel + vbQuestion, "Matches Known Profile?")
        If answer = vbCancel Then Exit Sub
        If answer = vbYes Then Me.Range("MatchesKnown" & Profile).Value = "Suspect"
        If answer = vbNo Then Me.Range("MatchesKnown" & Profile).Value = "None"
    End If

    Dim newCMF As cCODIS
    Set newCMF = Factory.CreateCODIS(Me.Range("Dest_CODIS" & Profile), SpecifyContributors)

End Sub

Private Sub Worksheet_Activate()

    If Application.WorksheetFunction.CountA(Range("Dest_CODIS1").Offset(0, 1)) = 0 Then Me.cmdExportCMF1.Visible = False
    If Application.WorksheetFunction.CountA(Range("Dest_CODIS2").Offset(0, 1)) = 0 Then Me.cmdExportCMF2.Visible = False
    
End Sub

Private Sub Worksheet_BeforeDelete()

    If AllSheets Is Nothing Or CODISSheets Is Nothing Then Admin.LoadSheetLists
    
    If AllSheets.Exists(Me.Name) Then AllSheets.Remove (Me.Name)
    If CODISSheets.Exists(Me.Name) Then CODISSheets.Remove (Me.Name)

End Sub
