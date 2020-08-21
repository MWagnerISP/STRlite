VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormImportGenTable 
   Caption         =   "Import Pre-STRmix or Standard from Genotype Table"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9225
   OleObjectBlob   =   "FormImportGenTable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormImportGenTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim counter As Integer
Dim fso As New FileSystemObject

Private Sub ImportGenTable_Click()

    GMID.Import_GenotypeTable
    
    If SampleNames Is Nothing Then 'if the file dialog was canceled and SampleNames wasn't created...
        Sheets("Master").Activate
        Admin.CleanUp
        Exit Sub
    End If
    
    FormImportGenTable.GMIDListBox.List = SampleNames.Keys
    FormImportGenTable.SampleListBox.Clear
    FormImportGenTable.StandardListBox.Clear

End Sub


Private Sub SampleAdd_Click()

    counter = 0
    
    For i = 0 To GMIDListBox.ListCount - 1
        If GMIDListBox.Selected(i - counter) Then
            SampleListBox.AddItem GMIDListBox.List(i - counter)
            GMIDListBox.RemoveItem (i - counter)
            counter = counter + 1
        End If
    Next i
    
End Sub

Private Sub SampleRemove_Click()

    counter = 0
    
    If SampleListBox.ListCount > 0 Then
        For i = 0 To SampleListBox.ListCount - 1
            If SampleListBox.Selected(i - counter) Then
                GMIDListBox.AddItem SampleListBox.List(i - counter)
                SampleListBox.RemoveItem (i - counter)
                counter = counter + 1
            End If
        Next i
    End If

End Sub

Private Sub SamplesDone_Click()

Application.ScreenUpdating = False

    If SampleListBox.ListCount = 0 And StandardListBox.ListCount = 0 Then
        Unload FormImportGenTable
        Exit Sub
    End If

    For i = 0 To SampleListBox.ListCount - 1
        If Not SelectedSamples.Exists(SampleListBox.List(i)) Then SelectedSamples.Add SampleListBox.List(i), Nothing
    Next i
    
    For i = 0 To StandardListBox.ListCount - 1
        If Not SelectedStandards.Exists(StandardListBox.List(i)) Then SelectedStandards.Add StandardListBox.List(i), Nothing
    Next i
    
    Application.StatusBar = "Creating allele summary worksheets..."
    
    Call GMID.HarvestSelectedSamples
    Call GMID.DumpSelectedSamples
    Call GMID.DumpSelectedStandards
    
    Application.StatusBar = False
    
    Admin.CleanUp
    
    Unload FormImportGenTable
    
Application.ScreenUpdating = True
End Sub

Private Sub StandardAdd_Click()
    
'    If StandardListBox.ListCount = 54 Then
'        MsgBox "Too many standards at once. Import these before you go back for more.", vbExclamation + vbOKOnly, "Easy there, champ."
'        Exit Sub
'    End If

    counter = 0
    
    For i = 0 To GMIDListBox.ListCount - 1
        If GMIDListBox.Selected(i - counter) Then
            StandardListBox.AddItem GMIDListBox.List(i - counter)
            GMIDListBox.RemoveItem (i - counter)
            counter = counter + 1
        End If
    Next i
    
End Sub

Private Sub StandardRemove_Click()

    counter = 0
    
    If StandardListBox.ListCount > 0 Then
        For i = 0 To StandardListBox.ListCount - 1
            If StandardListBox.Selected(i - counter) Then
                GMIDListBox.AddItem StandardListBox.List(i - counter)
                StandardListBox.RemoveItem (i - counter)
                counter = counter + 1
            End If
        Next i
    End If

End Sub


Private Sub UserForm_Initialize()

    Me.StandardListBox.Clear
    Me.SampleListBox.Clear
    Me.GMIDListBox.Clear
    Me.tbStdPageName.Text = ""
    
    Set SelectedSamples = Nothing
    Set SelectedSamples = New Scripting.Dictionary
    Set SelectedStandards = Nothing
    Set SelectedStandards = New Scripting.Dictionary

End Sub


Private Sub UserForm_Terminate()
    Admin.CleanUp
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub
