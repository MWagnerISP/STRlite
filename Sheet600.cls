VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Cmd_ClearStds_Click()
'Clear all standards

Dim proceed As Variant

Application.ScreenUpdating = False

    proceed = MsgBox("Are you sure you want to clear all standards on this sheet?", vbYesNo + vbQuestion, "Clear ALL Standards?")
    
    If proceed = vbNo Or proceed = vbCancel Then Exit Sub

    Me.Range("C8:H37").ClearContents
    Me.Range("N3").Value = "False" 'Uncheck checkboxes for controls
    
Application.ScreenUpdating = True

End Sub

Private Sub Cmd_ImportStds_Click()

    FormImportGenTable.Show
    
End Sub

Private Sub Worksheet_BeforeDelete()

    If AllSheets Is Nothing Or StandardSheets Is Nothing Then Admin.LoadSheetLists
    
    If AllSheets.Exists(Me.Name) Then AllSheets.Remove (Me.Name)
    If StandardSheets.Exists(Me.Name) Then StandardSheets.Remove (Me.Name)
    
End Sub
