Attribute VB_Name = "LR"
Dim LRCells As Collection 'list of the 6 cell addresses that LRs can go into (rngCorner)
'SelectedLRs key = LRfoldername, value = folder path

Dim availableLRSheets As Scripting.Dictionary 'key = Sheet name, item = EmptyLRs collection
Dim TotalEmptySpots As Integer
Option Explicit

'I really wish that VBA supported overloading subs....


Sub ImportSelectedLRs(SelectedLRs As Dictionary)
'For importing multiple LRs at once from the userform
'LRs are more complicated because a sheet holds multiple samples

    If SelectedLRs.Count = 0 Then Exit Sub
    If AllSheets Is Nothing Or LRSheets Is Nothing Then Call Admin.LoadSheetLists
    
    Set LRCells = New Collection
    LRCells.Add "D11"
    LRCells.Add "D24"
    LRCells.Add "D37"
    LRCells.Add "D58"
    LRCells.Add "D71"
    LRCells.Add "D84"

    Dim newLR As cLR
    Dim newLRSheet As Worksheet, lastLRsheet As Worksheet
    Dim EmptyLRs As Collection
    Dim newSheetName As String
    Dim i As Integer
    Dim newSheetCount As Integer: newSheetCount = 1
    Dim answer As Variant, LRfp As Variant, v As Variant
    
    TotalEmptySpots = 0
    
    Set availableLRSheets = New Scripting.Dictionary
    
    Call CheckForLRSpace 'fills availableLRsheets (if there are any) & counts TotalEmptySpots
    
    If availableLRSheets.Count > 0 And TotalEmptySpots > 0 Then 'If there are already available LR spots to use
        'Prompt for user choice on whether to add new LRs to existing sheets
        answer = MsgBox("Do you want to add these LRs to existing LR Worksheets if there's room?", vbYesNo, "Add to Existing Sheet?")
        
        If answer = vbYes Then
        
            For Each LRfp In SelectedLRs.Items 'LRfp = LR folder path
            
                'Make the cLR itself
                Set newLR = Factory.CreateLR(CStr(LRfp))

                If TotalEmptySpots > 0 Then
                    'Find the first spot to dump it
                    For Each v In availableLRSheets 'v is the sheet name, availableLRSheets(v) is the EmptyLRs collection
                        If availableLRSheets(v).Count > 0 Then 'If there's a spot on sheet v,
                        
                            'Dump the newLR to the target rng using .DumpLR
                            Call newLR.DumpLR(Worksheets(v), Range(availableLRSheets(v).Item(1)))
                            
                            availableLRSheets(v).Remove (1) 'remove the EmptyLR range that we just used
                            TotalEmptySpots = TotalEmptySpots - 1
                            Exit For
                        End If
                    Next v
                    
                    
                    
                Else: 'If there are no more empty spots, create a new sheet
                    
                    'Name the new sheet
                    If FormImportLR.tbLRPageName.Value <> "" Then
                        newSheetName = IIf(newSheetCount > 1, FormImportLR.tbLRPageName.Value & " (" & newSheetCount & ")", FormImportLR.tbLRPageName.Value)
                    Else: newSheetName = ""
                    End If
                
                    Set newLRSheet = Factory.CreateLRSheet(newSheetName)
                        If Not AllSheets.Exists(newLRSheet.Name) Then AllSheets.Add newLRSheet.Name, "LR"
                        If Not LRSheets.Exists(newLRSheet.Name) Then LRSheets.Add newLRSheet.Name, "LR"
                        
                    newSheetCount = newSheetCount + 1
                    
                    
                    Set EmptyLRs = New Collection
                    For Each v In LRCells 'all LR cells are available on this new sheet
                        EmptyLRs.Add v
                    Next v
                
                    Call newLR.DumpLR(newLRSheet, Range(EmptyLRs.Item(1)))
                    EmptyLRs.Remove (1)
                    
                    availableLRSheets.Add newLRSheet.Name, EmptyLRs
                    TotalEmptySpots = TotalEmptySpots + EmptyLRs.Count
                    
                End If
                
            Next LRfp
            Exit Sub
        End If
    End If


    'If User selects No or if there aren't any LR sheets yet, only use new sheets
    Set availableLRSheets = New Scripting.Dictionary 'refresh just in case
    TotalEmptySpots = 0
    
    'Make enough new sheets to hold SelectedLRs
    For i = 1 To CInt(Application.WorksheetFunction.RoundUp((SelectedLRs.Count / 6), 0))
    
        'Name the new sheet
        If FormImportLR.tbLRPageName.Value <> "" Then
            newSheetName = IIf(i > 1, FormImportLR.tbLRPageName.Value & " (" & i & ")", FormImportLR.tbLRPageName.Value)
        Else: newSheetName = ""
        End If
        
        'Make the new sheet
        Set newLRSheet = Factory.CreateLRSheet(newSheetName)
            If Not AllSheets.Exists(newLRSheet.Name) Then AllSheets.Add newLRSheet.Name, "LR"
            If Not LRSheets.Exists(newLRSheet.Name) Then LRSheets.Add newLRSheet.Name, "LR"
            
        Set EmptyLRs = New Collection
        For Each v In LRCells 'all LR cells are available on this new sheet
            EmptyLRs.Add v
        Next v
        
        availableLRSheets.Add newLRSheet.Name, EmptyLRs
        TotalEmptySpots = TotalEmptySpots + EmptyLRs.Count
    Next i
    
    

    For Each LRfp In SelectedLRs.Items 'LRfp = LR folder path
        
        'Make the cLR itself
        Set newLR = Factory.CreateLR(CStr(LRfp))
    
        'Find the first spot to dump it
        For Each v In availableLRSheets.Keys 'v is the sheet name, availableLRSheets(v) is the EmptyLRs collection
            If availableLRSheets(v).Count > 0 Then 'If there's a spot on sheet v,
                Call newLR.DumpLR(Worksheets(v), Range(availableLRSheets(v).Item(1))) 'use it
                availableLRSheets(v).Remove (1) 'and remove the spot from the list of EmptyLRs
                TotalEmptySpots = TotalEmptySpots - 1
                Set lastLRsheet = Worksheets(v)
                Exit For
            End If
        Next v
        
    Next LRfp
    
    lastLRsheet.Activate
                    
End Sub


Sub CheckForLRSpace()

    Dim ws As Worksheet, EmptyLRs As Collection, v As Variant
    
    'Check if there's already an LR Worksheet and if so, how much space is left on each
    For Each ws In Application.Worksheets
        If Left(ws.Name, 4) = "(LR)" Then
            
            'Detect whether LR spots are empty
                Set EmptyLRs = New Collection
                For Each v In LRCells
                    If ws.Range(v).Offset(0, 2).Value = "" Then EmptyLRs.Add v
                Next v
                
                availableLRSheets.Add ws.Name, EmptyLRs
                TotalEmptySpots = TotalEmptySpots + EmptyLRs.Count
        End If
    Next ws
    
End Sub


    
Sub Import_LRPrev(rngDest As Range)

    Call Admin.CrackTheHood("LR")

    Dim LRFolder As String
    
    'Prompt for LR folder
        With Application.FileDialog(msoFileDialogFolderPicker)
                .AllowMultiSelect = False
                .InitialFileName = Sheets("STRlite Settings").Range("STRmixResultsFolderpath").Value
                If .Show = -1 Then LRFolder = .SelectedItems(1)
        End With
        
    If LRFolder = "" Then Exit Sub
    
    Dim newLR As cLR
    Set newLR = Factory.CreateLR(LRFolder)
    
    Call newLR.DumpLR(rngDest.parent, rngDest)
    
    rngDest.parent.Activate

End Sub

Sub Import_LRPrevCombo(DestSheet As Worksheet, NOC As Integer)

    If DestSheet.Range("S2").Value = "" Then 'check for decon timestamp on DestSheet
        MsgBox "Import a deconvolution first.", vbOKOnly + vbInformation, "Need deconvolution!"
        Exit Sub
    End If
    
    Dim LRFolder As String
    
    'Prompt for LR folder
        With Application.FileDialog(msoFileDialogFolderPicker)
                .AllowMultiSelect = False
                .InitialFileName = Sheets("STRlite Settings").Range("STRmixResultsFolderpath").Value
                If .Show = -1 Then LRFolder = .SelectedItems(1)
        End With
        
    If LRFolder = "" Then Exit Sub
    
    Dim newLR As cLR
    Set newLR = Factory.CreateLR(LRFolder)
    
    If NOC = 1 Then newLR.DumpLR1P DestSheet
    If NOC = 2 Then newLR.DumpLR2P DestSheet
    
    DestSheet.Activate

End Sub

Sub ClearLR(rngCorner As Range)

        rngCorner.Offset(0, 2).ClearContents 'Sample ID
        rngCorner.Offset(1, 1).Resize(1, 4).ClearContents 'description/notes box
        rngCorner.Offset(6, 2).ClearContents 'Evidence File
        rngCorner.Offset(7, 2).ClearContents 'Timestamp
        rngCorner.Offset(8, 1).ClearContents 'Setting check
        rngCorner.Offset(8, 4).ClearContents 'NOC
        
        rngCorner.Offset(9, 1).ClearContents 'Omitted loci
        
        rngCorner.Offset(3, 14).Resize(1, 4).ClearContents 'HPD LR values
        rngCorner.Offset(7, 14).Resize(1, 4).ClearContents 'point LR values
        rngCorner.Offset(8, 16).Resize(2, 1).ClearContents 'strat/unified
        
        'Contributor section
        With rngCorner.Offset(3, 1).Resize(2, 4)
            .UnMerge
            .ClearContents
            .Font.Name = "Perpetua"
            .Font.Size = 12
            .Interior.Color = rngCorner.Offset(-1, 0).Interior.Color
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
        End With

End Sub

Public Function H1LR(rangeLR As Range)

    Select Case rangeLR.Value
    
        Case ""
            H1LR = ""
        
        Case Is = 0
            H1LR = rangeLR.Value
            
        Case Is > 1
            Dim sigfigs As Integer
            sigfigs = 3 - (1 + Fix(WorksheetFunction.log10(rangeLR.Value))) 'Fix returns only the integer portion (rounds down/truncates to integer)
            
            H1LR = Application.WorksheetFunction.RoundDown(rangeLR.Value, sigfigs)
            
        Case Else
            H1LR = ""
            
    End Select

End Function

Public Function H2LR(rangeLR As Range)

    Select Case rangeLR.Value
    
        Case ""
            H2LR = ""
                                                                        
        Case Is = 0
            H2LR = ""
                                                                        
        Case Is < 1
            H2LR = (1 / rangeLR.Value)
        
        Case Else
            H2LR = ""
        
    End Select

End Function


Public Function LRSummary(rangeLR As Range) As String

    If rangeLR.Cells(1, 1).Value = "" Then
        LRSummary = ""
        Exit Function
    End If

    Dim minLR As Double
    If rangeLR.Cells.Count > 1 Then
        'Dealing with population LRs. Take the minimum.
            minLR = Application.WorksheetFunction.Min(rangeLR)
    Else:
        minLR = rangeLR.Value
    End If
    
    'If minLR = 0 then you're done
    If minLR = 0 Then
        LRSummary = "The evidence does not support H1."
        Exit Function
    End If
    
    Dim H2 As Boolean: H2 = False
    
    If minLR < 1 Then
        minLR = 1 / minLR 'flip the LR
        H2 = True 'LR favors H2
    End If
    
    'For LR favoring H1 but <2 (1.999999 rounds conservatively down to LR = 1)
    If H2 = False And minLR < 2 Then
        LRSummary = "The evidence is equally likely if" & vbNewLine & "H1 or H2 are true. (LR = 1)"
        Exit Function
    End If
    
    
    Dim RawExponent As Integer: RawExponent = Int(Application.WorksheetFunction.Log(minLR, 10)) 'There is no VBA function for common log (base 10)
    Dim PowerAdjustment As Integer: PowerAdjustment = 3 * Int(RawExponent / 3)
    
    If NumberWords Is Nothing Then Call Admin.LoadSTRliteSettings
    Dim NumWord As String: NumWord = NumberWords(PowerAdjustment)
        NumWord = IIf(NumWord = "n/a", "", NumWord & " ") 'if NumWord= n/a (less than thousand), NumWord is empty. Otherwise add a space at the end.
    
    Dim NumBeforeWord As Double
    NumBeforeWord = minLR / (10 ^ PowerAdjustment)
    
    'Round conservatively; round up if favoring H2, round down if favoring H1. Two sig figs.
    Select Case NumBeforeWord
        Case Is >= 100
            NumBeforeWord = IIf(H2, Application.WorksheetFunction.RoundUp(NumBeforeWord, -1), Application.WorksheetFunction.RoundDown(NumBeforeWord, -1))
        Case Is >= 10
            NumBeforeWord = IIf(H2, Application.WorksheetFunction.RoundUp(NumBeforeWord, 0), Application.WorksheetFunction.RoundDown(NumBeforeWord, 0))
        Case Is < 10
            NumBeforeWord = IIf(H2, Application.WorksheetFunction.RoundUp(NumBeforeWord, 1), Application.WorksheetFunction.RoundDown(NumBeforeWord, 1))
    End Select
    
    
    'Put NumBeforeWord + NumWord together for FinalAnswer
    If H2 Then
        LRSummary = "The evidence is " & NumBeforeWord & " " & NumWord & "times more likely if H2 is true rather than H1."
    Else:
        LRSummary = "The evidence is " & NumBeforeWord & " " & NumWord & "times more likely if H1 is true rather than H2."
    End If

End Function
