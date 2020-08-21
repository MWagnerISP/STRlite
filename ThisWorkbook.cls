VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
'do NOT invalidate the ribbon during Workbook_Open! Breaks everything!


    'Ok because of different versions of Excel referencing different libraries
    'We need to force the references to certain libraries. I don't know if this does anything about missing libraries.
    Call ForceReferences

    Dim ws As Worksheet
    Dim nm As Name
    
    ThisWorkbook.Unprotect
    
    Call Admin.LociCollection
    Call Admin.LoadSTRliteSettings
    Call Admin.LockAll
    
    Call Sheets("Master").RefreshMaster
    
    'Call Admin.FixStupidButtons
    
'******Hide templates
    For Each ws In ActiveWorkbook.Worksheets
        If InStr(1, ws.Name, "Template") > 0 Then ws.Visible = xlVeryHidden
    Next ws
    Sheets("Import").Visible = xlVeryHidden
    Sheets("NIST 2017").Visible = xlVeryHidden

'******Hide Names
    For Each nm In ActiveWorkbook.Names
        nm.Visible = False
    Next nm

'******Make sure default sheets are visible
    Sheets("Master").Visible = True
    Sheets("About").Visible = True
    Sheets("STRlite Settings").Visible = xlHidden
    
'****** Reset Print Areas!
    Call Admin.ResetPrintAreas
    
End Sub


Sub ForceReferences()

    For i = ThisWorkbook.VBProject.References.Count To 1 Step -1
        Set theRef = ThisWorkbook.VBProject.References.Item(i)
    
        'if reference is "Missing" then remove it to avoid error message
        If theRef.isbroken = True Then
            ThisWorkbook.VBProject.References.Remove theRef
        End If
    
        'Debug.Print theRef.Description & ";" & theRef.FullPath & ";" & theRef.isbroken & vbCr
    Next i

    Call Admin.AddReferences(Me)

End Sub


