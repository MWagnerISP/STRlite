VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormAdmin 
   Caption         =   "Admin Tools"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4800
   OleObjectBlob   =   "FormAdmin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdFixButtons_Click()
    
    Call Admin.FixStupidButtons
    
End Sub

Private Sub CmdRestoreAll_Click()

    Dim ws As Worksheet
    Dim nm As Name
 
Application.ScreenUpdating = False
 
'******Hide templates, Import, & NIST
    For Each ws In ActiveWorkbook.Worksheets
        If InStr(1, ws.Name, "Template") > 0 Then ws.Visible = xlVeryHidden
    Next ws
    
    Sheets("Import").Visible = xlVeryHidden
    Sheets("NIST 2017").Visible = xlVeryHidden
    
'******Hide Names
    'This hides the PrintArea range names, which erases set print areas on all tabs
    For Each nm In ActiveWorkbook.Names
        nm.Visible = False
    Next nm

'So we re-set the print areas
Call Admin.ResetPrintAreas

Call Admin.LockAll

Application.ScreenUpdating = True
    
End Sub

Private Sub CmdShowNames_Click()
    
    Dim nm As Name
    
    For Each nm In ActiveWorkbook.Names
        nm.Visible = True
    Next nm
    
    
End Sub

Private Sub CmdUnhideAll_Click()
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub CmdUnlockAdmin_Click()

    If TextBox_Password.Value = STRlitePW Then
        STRliteUnlocked = True
        CmdUnhideAll.Visible = True
        CmdUnlockAll.Visible = True
        CmdShowNames.Visible = True
        CmdRestoreAll.Visible = True
        CmdFixButtons.Visible = True
        TextBox_Password.Visible = False
        Label1.Caption = "Correct!"
        CmdUnlockAdmin.Visible = False
    Else
        Label1.Caption = "Incorrect. Try again:"
    End If
    
End Sub


Private Sub CmdUnlockAll_Click()
    
Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Unprotect (STRlitePW)
    Next ws

Application.ScreenUpdating = True
    'Sheets("Deconvolution").Chart_MixProp.Unprotect (STRlitePW)
    
End Sub

Private Sub UserForm_Initialize()

If STRliteUnlocked Then
        CmdUnhideAll.Visible = True
        CmdUnlockAll.Visible = True
        CmdShowNames.Visible = True
        CmdRestoreAll.Visible = True
        CmdFixButtons.Visible = True
        TextBox_Password.Visible = False
        CmdUnlockAdmin.Visible = False
        Label1.Caption = ""
        
    Else:
        CmdUnhideAll.Visible = False
        CmdUnlockAll.Visible = False
        CmdShowNames.Visible = False
        CmdRestoreAll.Visible = False
        CmdFixButtons.Visible = True
        TextBox_Password.Value = ""
End If

End Sub
