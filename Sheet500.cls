VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub cmdImportDecon_Click()

    Dim DeconPath As String
    'Prompt for Decon folder
        With Application.FileDialog(msoFileDialogFolderPicker)
                .InitialFileName = Sheets("STRlite Settings").Range("STRmixResultsFolderPath").Value
                If .Show = -1 Then
                    DeconPath = .SelectedItems(1)
                    Call ImportDecon(DeconPath, Me, "2P")
                End If
        End With

End Sub

Private Sub cmdImportLR_Click()

    Call LR.Import_LRPrevCombo(Me, 2)
    Me.Select
    
End Sub

Private Sub CondtoCODIS1_Click()

    Dim answer

    answer = MsgBox("Are you sure? This is a conditioned profile." & vbNewLine & _
    "You should almost never need to submit this to CODIS." _
    & vbNewLine & vbNewLine & _
    "Proceed?", vbYesNo + vbQuestion, "Submit Conditioned Profile?")

    If answer = vbYes Then Call Me.SendToCODIS(1, True)
    If answer = vbNo Then Exit Sub
    
End Sub

Private Sub ToCODIS1_Click()
    Call Me.SendToCODIS(1, False)
End Sub

Private Sub ToCODIS2_Click()
    Call Me.SendToCODIS(2, False)
End Sub

Sub SendToCODIS(cont As Integer, Cond As Boolean)

    'Detect CODIS sheet or create new one if needed:
    If PubFun.WorksheetExists(Replace(Me.Name, "(2P)", "(C)", 1)) Then
        Set TargetCODISsheet = Sheets(Replace(Me.Name, "(2P)", "(C)", 1))
    Else: Set TargetCODISsheet = Factory.CreateCODISSheet(Me, "2P")
    End If
    
    'Ranges on Decon sheet:
    Dim rngContributor As Range: Set rngContributor = Me.Range("Dest_2PComponents").Offset(0, cont)
    Dim rngCategories As Range: Set rngCategories = Me.Range("Dest_2PLocusCat").Offset(0, cont)
    Dim rngWholeCategory As Range: Set rngWholeCategory = Me.Range("Dest_2PSpecCat").Offset(0, cont)
    
    'Ranges of actual alleles on CODIS sheet:
    Dim rngCODIS1 As Range: Set rngCODIS1 = TargetCODISsheet.Range("Dest_CODIS1").Offset(0, 1)
    Dim rngCODIS2 As Range: Set rngCODIS2 = TargetCODISsheet.Range("Dest_CODIS2").Offset(0, 1)
    
    Dim CODISfull1 As Boolean, CODISfull2 As Boolean
    
'Determine if CODIS profiles are already present (CODISfull)
        
    CODISfull1 = IIf(Application.WorksheetFunction.CountA(rngCODIS1) > 0, True, False)
    CODISfull2 = IIf(Application.WorksheetFunction.CountA(rngCODIS2) > 0, True, False)
    
'If both profiles are full then make a fresh sheet
    If CODISfull1 And CODISfull2 Then
        Set TargetCODISsheet = Factory.CreateCODISSheet(Me, "2P") 'create new sheet
        Set rngCODIS1 = TargetCODISsheet.Range("Dest_CODIS1").Offset(0, 1) 'reset the target ranges to the new sheet
        Set rngCODIS2 = TargetCODISsheet.Range("Dest_CODIS2").Offset(0, 1)
        CODISfull1 = False
        CODISfull2 = False
    End If
        
'Transfer profile to first available CODIS slot. The "transferred" ranges must be the same size to do it this way. No merged cells.
    Dim i As String
    i = IIf(CODISfull1, "2", "1")
    
    TargetCODISsheet.Range("Dest_CODIS" & i).Offset(0, 1).Value = rngContributor.Value 'alleles
    TargetCODISsheet.Range("Dest_CODIS" & i).Offset(0, 2).Value = rngCategories.Value 'locus categories ("M")
    TargetCODISsheet.Range("SendtoCODIS" & i).Offset(-1, -1).Value = cont 'saves the contributor number from the decon
    TargetCODISsheet.Range("SendtoCODIS" & i).Offset(-2, 0).Value = rngWholeCategory.Value 'specimen category
    TargetCODISsheet.Range("MatchesKnown" & i).Offset(2, 0).Value = rngWholeCategory.Offset(5, 0).Value 'MME
        If rngWholeCategory.Offset(5, 0).Value < CODISgoalMME Then TargetCODISsheet.Range("SendtoCODIS" & i).Offset(-2, 0).Value = "Targeted"
    TargetCODISsheet.Range("MatchesKnown" & i).Offset(3, 0).Value = rngWholeCategory.Offset(6, 0).Value 'SDIS MRE
    TargetCODISsheet.Range("MatchesKnown" & i).Offset(4, 0).Value = rngWholeCategory.Offset(7, 0).Value 'NDIS MRE
    If Conditioned = True Then TargetCODISsheet.Range("MatchesKnown" & i).Value = "Conditioned Suspect"

    If i = "1" Then TargetCODISsheet.cmdExportCMF1.Visible = True
    If i = "2" Then TargetCODISsheet.cmdExportCMF2.Visible = True
    
TargetCODISsheet.Select

End Sub


Private Sub Worksheet_Activate()
    On Error Resume Next
    Me.ToCODIS1.Visible = IIf(Application.WorksheetFunction.CountA(Me.Range("Dest_2PLociPreSTRmix").Offset(0, 4)) = 0, False, True)
    Me.ToCODIS2.Visible = IIf(Application.WorksheetFunction.CountA(Me.Range("Dest_2PLociPreSTRmix").Offset(0, 5)) = 0, False, True)
    
    Me.CondtoCODIS1.Visible = IIf(Me.Range("S51").Value > 0, True, False)
    On Error GoTo 0
End Sub

Private Sub Worksheet_BeforeDelete()

    If AllSheets Is Nothing Or DoubleSheets Is Nothing Then Admin.LoadSheetLists
    
    If AllSheets.Exists(Me.Name) Then AllSheets.Remove (Me.Name)
    If DoubleSheets.Exists(Me.Name) Then DoubleSheets.Remove (Me.Name)
    
End Sub

