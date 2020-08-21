VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormImportDecons 
   Caption         =   "Import Multiple Deconvolutions"
   ClientHeight    =   10770
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9915
   OleObjectBlob   =   "FormImportDecons.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormImportDecons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ParentPath As String
Dim SubFolders As Scripting.Dictionary
Dim ChosenFolders As New Scripting.Dictionary 'key = folder name, value = folder path
Dim fso As New FileSystemObject


Private Sub chkOmitDBSearch_Click()
    Call ListSubfolders(ParentPath)
End Sub

Private Sub chkOmitLRPrev_Change()
    Call ListSubfolders(ParentPath)
End Sub


Private Sub cmdFolderDown_Click()

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


Private Sub lbFolders_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.lbFolders.ListIndex = -1 Then Exit Sub
    
    Dim NewPath As String
    NewPath = SubFolders(Me.lbFolders.List(Me.lbFolders.ListIndex))
    
    Call ListSubfolders(NewPath)
    ParentPath = NewPath

End Sub

Private Sub UserForm_Initialize()

    Me.cmdFolderUp.Caption = ChrW(708) & vbNewLine & "Folder Up"
    Me.cmdFolderDown.Caption = "Folder Down" & vbNewLine & ChrW(709)
    
    ParentPath = Sheets("STRlite Settings").Range("STRmixResultsFolderpath").Value
    
    'If STRlite Settings doesn't have a valid path then just use STRlite's location
    If Not fso.FolderExists(ParentPath) Then ParentPath = ThisWorkbook.Path
    
    Call ListSubfolders(ParentPath)

End Sub


Private Sub cmdChangeParentFolder_Click()
    
    'Prompt for parent folder
        With Application.FileDialog(msoFileDialogFolderPicker)
                .InitialFileName = ParentPath
                If .Show = -1 Then ParentPath = .SelectedItems(1)
        End With
    
    Call ListSubfolders(ParentPath)

End Sub

Private Sub cmdAddDecon_Click()

    If lbChosenDeconFolders.ListCount > 12 Then
        MsgBox "Too many decons at once. Import these before you go back for more.", vbExclamation + vbOKOnly, "Easy there, champ."
        Exit Sub
    End If

    Dim str As String, i As Integer
    Dim counter As Integer: counter = 0
    
    For i = 0 To lbFolders.ListCount - 1
        If lbFolders.Selected(i - counter) Then
            str = lbFolders.List(i - counter)

            lbChosenDeconFolders.AddItem str
            lbFolders.RemoveItem (i - counter)
            If Not ChosenFolders.Exists(str) Then ChosenFolders.Add str, SubFolders(str)
            counter = counter + 1
            
        End If
    Next i
    
End Sub

Private Sub cmdRemoveDecon_Click()

    Dim counter As Integer: counter = 0
    Dim str As String
    
    If lbChosenDeconFolders.ListCount > 0 Then
        For i = 0 To lbChosenDeconFolders.ListCount - 1
            If lbChosenDeconFolders.Selected(i - counter) Then
                str = lbChosenDeconFolders.List(i - counter)
                lbFolders.AddItem str
                lbChosenDeconFolders.RemoveItem (i - counter)
                If ChosenFolders.Exists(str) Then ChosenFolders.Remove (str)
                counter = counter + 1
            End If
        Next i
    End If

End Sub

Private Sub cmdImportDecons_Click()

    Application.ScreenUpdating = False
    Dim dFolder As Variant
    
    For Each dFolder In ChosenFolders.Items
        Call Decon.ImportDecon(CStr(dFolder), , "Decon")
    Next dFolder
    
    Admin.CleanUp
    
    Application.ScreenUpdating = True
    Unload Me

End Sub


Private Sub ListSubfolders(ParentPath As String)

    Dim fso As New Scripting.FileSystemObject
    
    Dim fsoParent As Object: Set fsoParent = fso.GetFolder(ParentPath)
    Dim fsoSubFolder As Folder

    Set SubFolders = New Scripting.Dictionary
    
    For Each fsoSubFolder In fsoParent.SubFolders
        SubFolders.Add fsoSubFolder.Name, fsoSubFolder.Path
    Next
    
    If Me.chkOmitDBSearch.Value = True Then
        For Each v In SubFolders.Keys
            If InStr(1, v, "DBSearch") > 0 Then SubFolders.Remove v
        Next v
    End If
    
    If Me.chkOmitLRPrev.Value = True Then
        For Each v In SubFolders.Keys
            If InStr(1, v, "LRPrev") > 0 Then SubFolders.Remove v
        Next v
    End If
    
    Me.lbFolders.List = SubFolders.Keys
    
End Sub


Private Sub UserForm_Terminate()
    Admin.CleanUp
    Set fso = Nothing
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub
