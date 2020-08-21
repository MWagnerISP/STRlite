VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormImportLR 
   Caption         =   "Import/Organize LRs"
   ClientHeight    =   10770
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9915
   OleObjectBlob   =   "FormImportLR.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormImportLR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ParentPath As String
Dim SubFolders As Scripting.Dictionary
Dim ChosenFolders As New Scripting.Dictionary 'key = folder name, value = folder path
Dim fso As New FileSystemObject

Private Sub UserForm_Initialize()

    Me.cmdFolderUp.Caption = ChrW(708) & vbNewLine & "Folder Up"
    Me.cmdFolderDown.Caption = "Folder Down" & vbNewLine & ChrW(709)

    ParentPath = Sheets("STRlite Settings").Range("STRmixResultsFolderpath").Value
    
    'If STRlite Settings doesn't have a valid path then just use STRlite's location
    If Not fso.FolderExists(ParentPath) Then ParentPath = ThisWorkbook.Path
    
    Me.chkLROnly.Value = False
    
    'Call ListSubfolders(ParentPath)
    
End Sub


Private Sub cmdDone_Click()
    If Me.lbChosenLRFolders.ListCount > 0 Then
        Dim answer As Variant
        answer = MsgBox("Are you done importing LRs?", vbYesNo + vbQuestion, "Done Importing?")
        If answer = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdImportLRs_Click()

Application.ScreenUpdating = False

    Call LR.ImportSelectedLRs(ChosenFolders)
    
    ChosenFolders.RemoveAll
    Set ChosenFolders = New Dictionary
    
    Me.lbChosenLRFolders.Clear
    Me.tbLRPageName.Value = ""
    
    MsgBox "LR import complete!", vbOKOnly, "All Done!"
    
Application.ScreenUpdating = True

End Sub


Private Sub cmdChangeParentFolder_Click()

    'Prompt for parent folder
        With Application.FileDialog(msoFileDialogFolderPicker)
                .InitialFileName = ParentPath
                If .Show = -1 Then ParentPath = .SelectedItems(1)
        End With
    
    Call ListSubfolders(ParentPath)
    
End Sub

Private Sub cmdFolderDown_Click()

    If Me.lbFolders.ListIndex = -1 Then Exit Sub
    
    Dim NewPath As String
    NewPath = SubFolders(Me.lbFolders.List(Me.lbFolders.ListIndex))
    
    Call ListSubfolders(NewPath)
    ParentPath = NewPath

End Sub

Private Sub lbFolders_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Same as "Folder Down"

    If Me.lbFolders.ListIndex = -1 Then Exit Sub
    
    Dim NewPath As String
    NewPath = SubFolders(Me.lbFolders.List(Me.lbFolders.ListIndex))
    
    Call ListSubfolders(NewPath)
    ParentPath = NewPath

End Sub

Private Sub cmdFolderUp_Click()

    'If ParentPath ends in \, it is the root drive and not a folder. Can't go any further up.
    If Right(ParentPath, 1) = "\" Then Exit Sub

    Dim fso As New Scripting.FileSystemObject
    Dim fsoUp As Object: Set fsoUp = fso.GetFolder(ParentPath).ParentFolder
    
    Call ListSubfolders(fsoUp.Path)
    ParentPath = fsoUp.Path
    
End Sub


Private Sub cmdAddLR_Click()

    If lbChosenLRFolders.ListCount = 18 Then
        MsgBox "Too many LRs at once. Import these before you go back for more.", vbExclamation + vbOKOnly, "Easy there, champ."
        Exit Sub
    End If
    
    Dim str As String, i As Integer
    Dim counter As Integer: counter = 0
    
    For i = 0 To lbFolders.ListCount - 1
        If lbFolders.Selected(i - counter) Then
            str = lbFolders.List(i - counter)

            'Check folder for correct files
            If CheckLRFolder(SubFolders(str)) Then
                lbChosenLRFolders.AddItem str
                lbFolders.RemoveItem (i - counter)
                If Not ChosenFolders.Exists(str) Then ChosenFolders.Add str, SubFolders(str)
                counter = counter + 1
            End If
            
        End If
    Next i
    
End Sub


Private Sub cmdRemoveLR_Click()

    Dim counter As Integer: counter = 0
    Dim str As String, i As Integer
    
    If lbChosenLRFolders.ListCount > 0 Then
        For i = 0 To lbChosenLRFolders.ListCount - 1
            If lbChosenLRFolders.Selected(i - counter) Then
                str = lbChosenLRFolders.List(i - counter)
                lbFolders.AddItem str
                lbChosenLRFolders.RemoveItem (i - counter)
                If ChosenFolders.Exists(str) Then ChosenFolders.Remove (str)
                counter = counter + 1
            End If
        Next i
    End If
    
End Sub


Private Sub chkLROnly_Change()
    Call ListSubfolders(ParentPath)
End Sub


Private Sub ListSubfolders(ParentPath As String)

    Dim fso As New Scripting.FileSystemObject
    
    Dim fsoParent As Object: Set fsoParent = fso.GetFolder(ParentPath)
    Dim fsoSubFolder As Folder
    
    'Refresh SubFolders every time we run this sub:
    Set SubFolders = New Scripting.Dictionary
    
    For Each fsoSubFolder In fsoParent.SubFolders
        SubFolders.Add fsoSubFolder.Name, fsoSubFolder.Path
    Next
    
    If Me.chkLROnly.Value = True Then
        For Each v In SubFolders.Keys
            If InStr(1, v, "LR") = 0 Then SubFolders.Remove v
        Next v
    End If
    
    Me.lbFolders.List = SubFolders.Keys

End Sub

Private Function CheckLRFolder(LRFolder As String) As Boolean
'Checks if required files are in the selected folder and makes sure they're the right kind
    
    If Not fso.FileExists(LRFolder & "/results.xml") Or Not fso.FileExists(LRFolder & "/config.xml") Then
        CheckLRFolder = False
        MsgBox "Error validating LR folder: " & vbNewLine & vbNewLine & LRFolder & vbNewLine & vbNewLine & _
                "Folder does not contain the required files.", vbCritical + vbOKOnly, "Wrong Folder?"
        Exit Function
    End If

    With CreateObject("MSXML2.DOMDocument")
        .async = False
        .validateOnParse = False
        .Load (LRFolder & "/config.xml")
        
        'The tag <lrSettings> indicates an LR config file
        If .selectSingleNode("//lrSettings") Is Nothing Then
            MsgBox "Error validating LR folder: " & vbNewLine & vbNewLine & LRFolder & vbNewLine & vbNewLine & _
                "Config.xml file is not from an LR.", vbCritical + vbOKOnly, "Wrong file!"
            CheckLRFolder = False
            Exit Function
        End If
    End With
    
    With CreateObject("MSXML2.DOMDocument")
        .async = False
        .validateOnParse = False
        .Load (LRFolder & "/results.xml")
        
        'The tag <lrSummary> indicates an LR results file
        If .selectSingleNode("//lrSummary") Is Nothing Then
            MsgBox "Error validating LR folder: " & vbNewLine & vbNewLine & LRFolder & vbNewLine & vbNewLine & _
                "Results.xml file is not from an LR.", vbCritical + vbOKOnly, "Wrong file!"
            CheckLRFolder = False
            Exit Function
        End If
    End With
    
    CheckLRFolder = True
    
End Function



Private Sub UserForm_Terminate()
    Admin.CleanUp
    Set fso = Nothing
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub
