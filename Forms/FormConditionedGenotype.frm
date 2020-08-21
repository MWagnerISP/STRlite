VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormConditionedGenotype 
   Caption         =   "Select Standard to Condition On"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   OleObjectBlob   =   "FormConditionedGenotype.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormConditionedGenotype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public wsDest As Worksheet

Private Sub Cmd_Browse_Click()

    Application.ScreenUpdating = False

    GMID.Import_GenotypeTable
    
    wsDest.Activate

    '******Put sample names in the list box
    Me.List_SampleList.Clear
    Me.List_SampleList.List = SampleNames.Keys
    
    Application.ScreenUpdating = True
        
End Sub

Private Sub Cmd_Std1_Click()

    If Me.List_SampleList.ListCount = 0 Then
        MsgBox "Please browse for the Genotype Table file containing your standard.", vbOKOnly + vbInformation, "If you seek nothing..."
        Exit Sub
    End If
    
    If Me.List_SampleList.ListIndex = -1 Then
        MsgBox "Please select a standard to import.", vbOKOnly + vbInformation, "If you seek nothing..."
        Exit Sub
    End If

Application.ScreenUpdating = False

    For i = 0 To List_SampleList.ListCount - 1
        If List_SampleList.Selected(i) Then CondStd = CStr(List_SampleList.List(i))
    Next i
    
    wsDest.Range("E:F").EntireColumn.Hidden = False
    
    Call GMID.HarvestDumpConditionedStandard(CondStd, wsDest)
    
    wsDest.OLEObjects("cmdHideCond").Visible = True
    
    Admin.CleanUp
    
    Unload Me
    
Application.ScreenUpdating = True
    
End Sub


Private Sub UserForm_Terminate()

    Sheets("Import").Visible = xlVeryHidden
    wsDest.Activate
    Unload Me
    Application.ScreenUpdating = True
    
End Sub
