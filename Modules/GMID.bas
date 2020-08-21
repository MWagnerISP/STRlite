Attribute VB_Name = "GMID"
Option Explicit
Public GMIDSample As String
Public PropCount As Integer
Public Contributors As Integer

Public SampleNames As Scripting.Dictionary
Public SelectedSamples As Scripting.Dictionary
Public SelectedStandards As Scripting.Dictionary

Public CondStd As String
Public newCondStd As cProfileGenotype

Dim SampleObjects As Scripting.Dictionary 'key = SampleName, item = cProfileGenotype
Dim StandardObjects As Scripting.Dictionary 'key = SampleName, item = cProfileGenotype

Dim Overwrite As Boolean


Sub Import_GenotypeTable()

    Call Admin.CrackTheHood("GMID")
    Call Admin.LoadSTRliteSettings
    
    Application.ScreenUpdating = False

'******Variables*****

    Dim wkbTemp As Workbook
    
    Dim CurrentFile As String
    Dim FilePath_GenTable As String
    Dim FileName_GenTable As String
    Dim fso As New FileSystemObject
    Dim StartPath As String

    Dim Samples As Range
    
    Dim rng As Range
    
    CurrentFile = ThisWorkbook.Name
    
    StartPath = Sheets("STRlite Settings").Range("GenotypeTable_FolderPath").Value
    
    'If STRlite Settings doesn't have a valid path then just use STRlite's location
    If Not fso.FolderExists(StartPath) Then StartPath = ThisWorkbook.Path
    
'******Prompt for Genotype Table File******
    With Application.FileDialog(msoFileDialogFilePicker)
            .AllowMultiSelect = False
            .InitialFileName = StartPath
            .Title = "Open Genotype Table file (.txt)"
            .Filters.Add "Text", "*.txt", 1
            If .Show = -1 Then FileName_GenTable = .SelectedItems(1)
    End With

    If FileName_GenTable = "" Then Exit Sub
    
    Sheets("Import").Cells.ClearContents
    
    Workbooks.OpenText Filename:=FileName_GenTable, DataType:=xlDelimited, Tab:=True
    
    Set wkbTemp = ActiveWorkbook
    
    FilePath_GenTable = wkbTemp.Path
    
    If wkbTemp.Sheets(1).Range("B1").Value <> "Marker" Or InStr(1, wkbTemp.Sheets(1).Range("A1").Value, "Sample", vbBinaryCompare) = 0 Then
        MsgBox "This file does not contain a correctly formatted Genotype Table.", vbCritical, "Try Again"
        wkbTemp.Close SaveChanges:=False
        Call Admin.CleanUp
        Exit Sub
    End If
    
    
    '******Copy genotype table to CurrentFile and close wkbtemp
    
    wkbTemp.Sheets(1).Range(Cells(1, 1), Cells(5000, (NumGenotypeTableAlleles * 3) + 2)).Select
    Selection.Copy
    Windows(CurrentFile).Activate
    'Sheets("Import").Activate
    
    Sheets("Import").Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    wkbTemp.Close SaveChanges:=False
    Set wkbTemp = Nothing
    
    
    '******Populate SampleName dictionary with unique values
    Windows(CurrentFile).Activate
    Set SampleNames = New Scripting.Dictionary
    Application.ScreenUpdating = False
    Sheets("Import").Activate
    Set Samples = Sheets("Import").Range(Cells(1, 1), Cells(Rows.Count, 1).End(xlUp))

    Dim v As Variant, e As Variant

    v = Samples.Value 'v = array of cell values

    With SampleNames
        .CompareMode = 1
        For Each e In v
            'Don't include RBs, ladders, and PCs in the list (Amp blanks/ABs are harder to exclude based on text)
            If InStr(1, e, "RB", vbTextCompare) + InStr(1, e, "ladder", vbTextCompare) + InStr(1, e, "PC", vbBinaryCompare) = 0 Then
                If Not .Exists(e) Then .Add e, Nothing
            End If
        Next
    End With

    On Error Resume Next
        SampleNames.Remove ("Sample Name")
        SampleNames.Remove ("Sample File")
    On Error GoTo 0
    
End Sub

Sub HarvestSelectedSamples()
'We're gonna do both unknowns and standards at the same time.
    
    'Samples = every used cell in column A, starting with row 2
    Dim Samples As Range: Set Samples = Sheets("Import").Range(Cells(2, 1), Cells(Rows.Count, 1).End(xlUp))
    
    Set SampleObjects = New Scripting.Dictionary
    Set StandardObjects = New Scripting.Dictionary
    
    Dim rng As Range
    Dim newSample As cProfileGenotype
    Dim newLocus As cLocusWhole
    
    'To speed things up, I only want to iterate through the actual Genotype Table ONCE.

    For Each rng In Samples
        
        'If the sample is one of the SelectedSamples,
        If SelectedSamples.Exists(rng.Value) Then
        
            'if we haven't made the cProfileGenotype yet, make it:
            If Not SampleObjects.Exists(rng.Value) Then
                Set newSample = New cProfileGenotype
                newSample.SampleName = rng.Value
                newSample.IsStandard = False
                SampleObjects.Add rng.Value, newSample
            End If
            
            'now we have a cProfileGenotype object stored in SampleObjects
            'Create a new locus from this row
            Set newLocus = CreateLocusGMID(rng.Offset(0, 1), newSample.IsStandard)
            SampleObjects(rng.Value).Loci.Add newLocus.LocusName, newLocus 'add the locus to the corresponding cProfileGenotype.Loci
            
            If newLocus.SaturationFlag Then MsgBox "Peak(s) above " & SaturationMax & " RFU detected in " & newLocus.LocusName & " of " & rng.Value & "." & vbNewLine & _
            "Please double-check your electropherogram.", vbExclamation, "Saturation Detected!"
            
        End If
        
        If SelectedStandards.Exists(rng.Value) Then
        
            'if we haven't made the cProfileGenotype yet, make it:
            If Not StandardObjects.Exists(rng.Value) Then
                Set newSample = New cProfileGenotype
                newSample.SampleName = rng.Value
                newSample.IsStandard = True
                StandardObjects.Add rng.Value, newSample
            End If
            
            'now we have a cProfileGenotype object stored in StandardObjects
            'Create a new locus from this row
            Set newLocus = CreateLocusGMID(rng.Offset(0, 1), newSample.IsStandard)
            StandardObjects(rng.Value).Loci.Add newLocus.LocusName, newLocus 'add the locus to the corresponding cProfileGenotype.Loci
            
            If newLocus.SaturationFlag Then MsgBox "Peak(s) above " & SaturationMax & " RFU detected in " & newLocus.LocusName & " of " & rng.Value & "." & vbNewLine & _
            "Please double-check your electropherogram.", vbExclamation, "Saturation Detected!"
            
        End If
        
    Next rng

                
End Sub


Sub DumpSelectedSamples()

    'Samples & Standards are now stored in SampleObjects & StandardObjects dictionaries
    Dim newGMIDSheet As Worksheet, answer As Variant, v As Variant
    
    'Samples/unknowns go on the Pre-STRmix sheet
    For Each v In SampleObjects.Keys
        'Figure out which worksheet to put the sample on:
        If PubFun.WorksheetExists(Left("(P) " & PubFun.FixWorksheetName(CStr(v)), 31)) Then
            answer = MsgBox("You already have a Pre-STRmix worksheet for that sample." & vbNewLine & vbNewLine & _
                        "Do you want to overwrite it?", vbYesNo, "Overwrite Existing Sample?")
                        If answer = vbNo Or answer = vbCancel Then GoTo SkipThisOne
                        If answer = vbYes Then Set newGMIDSheet = Sheets(Left("(P) " & PubFun.FixWorksheetName(CStr(v)), 31))
        Else:
            Set newGMIDSheet = Factory.CreateGMIDSheet(SampleObjects(v))
        End If
        
        'Then dump the sample
        Call SampleObjects(v).DumpData(newGMIDSheet)
        
        If AllSheets Is Nothing Or GMIDSheets Is Nothing Then Call Admin.LoadSheetLists
        If Not AllSheets.Exists(newGMIDSheet.Name) Then AllSheets.Add newGMIDSheet.Name, "PreSTRmix"
        If Not GMIDSheets.Exists(newGMIDSheet.Name) Then GMIDSheets.Add newGMIDSheet.Name, "PreSTRmix"
        
SkipThisOne:
    Next v

End Sub

Sub DumpSelectedStandards()
'Standards are a bit more complicated because a sheet holds multiple samples

    If SelectedStandards.Count = 0 Then Exit Sub
    If AllSheets Is Nothing Or StandardSheets Is Nothing Then Call Admin.LoadSheetLists

    Dim newStdSheet As Worksheet, ws As Worksheet
    Dim i As Integer
    Dim answer As Variant, std As Variant, v As Variant
    Dim StdSheets As New Scripting.Dictionary 'key = Sheet name, item = EmptyCol collection
    Dim EmptyCol As Collection, tempColl As Collection 'contains empty column position(s) (1 to 6)
    Dim TotalEmptySpots As Integer: TotalEmptySpots = 0
    
    'Check if there's already a Standards Worksheet and if so, how much space is left on each
    For Each ws In Application.Worksheets
        If Left(ws.Name, 4) = "(Std" Then
            
            'Detect whether columns are empty
                Set EmptyCol = New Collection
                For i = 1 To 6
                    If ws.Range("Dest_StandardSampleName").Offset(0, i) = "" Then EmptyCol.Add i
                Next i
                
                StdSheets.Add ws.Name, EmptyCol
                TotalEmptySpots = TotalEmptySpots + EmptyCol.Count
        End If
    Next ws
                
    If StdSheets.Count And TotalEmptySpots > 0 Then
        'Prompt for user choice on whether to add SelectedStandards to existing sheets
        answer = MsgBox("Do you want to add these standards to existing Standard Worksheets if there's room?", vbYesNo, "Add to Existing Sheet?")
        
        If answer = vbYes Then
            For Each std In SelectedStandards.Keys

                If TotalEmptySpots > 0 Then
                    'Find the first spot to dump it
                    For Each v In StdSheets 'v is the sheet name, StdSheets(v) is the EmptyCol collection
                        Set tempColl = StdSheets(v)
                        If tempColl.Count > 0 Then 'If there's a spot on sheet v,
                            Call StandardObjects(std).DumpData(Worksheets(v), False, tempColl.Item(1))
                            tempColl.Remove (1)
                            TotalEmptySpots = TotalEmptySpots - 1
                            Exit For
                        End If
                    Next v
                    
                Else: 'If there are no more empty spots, create a new sheet
                    Set newStdSheet = Factory.CreateStandardSheet(FormImportGenTable.tbStdPageName.Value)
                        If Not AllSheets.Exists(newStdSheet.Name) Then AllSheets.Add newStdSheet.Name, "Standard"
                        If Not StandardSheets.Exists(newStdSheet.Name) Then StandardSheets.Add newStdSheet.Name, "Standard"
                    Set EmptyCol = New Collection
                        For i = 2 To 6 'we're going to use the first column for this std, so 2 to 6 will be empty
                            EmptyCol.Add i
                        Next i
                    StdSheets.Add newStdSheet.Name, EmptyCol
                    TotalEmptySpots = TotalEmptySpots + 5
                    
                    Call StandardObjects(std).DumpData(newStdSheet, False, 1) 'use the first column of this new sheet
                    
                End If
                
            Next std
            Exit Sub
        End If
    End If

    'If User selects No or if there aren't any Standard sheets yet, only use new sheets
        Set StdSheets = New Scripting.Dictionary
        TotalEmptySpots = 0
        Dim newSheetName As String
        
        'Make enough new sheets to hold SelectedStandards
        For i = 1 To CInt(Application.WorksheetFunction.RoundUp((SelectedStandards.Count / 6), 0))
            'Name the new sheet
            If FormImportGenTable.tbStdPageName.Value <> "" Then
                newSheetName = IIf(i > 1, FormImportGenTable.tbStdPageName.Value & " (" & i & ")", FormImportGenTable.tbStdPageName.Value)
            Else: newSheetName = ""
            End If
            
            'Make the new sheet
            Set newStdSheet = Factory.CreateStandardSheet(newSheetName)
                If Not AllSheets.Exists(newStdSheet.Name) Then AllSheets.Add newStdSheet.Name, "Standard"
                If Not StandardSheets.Exists(newStdSheet.Name) Then StandardSheets.Add newStdSheet.Name, "Standard"
            Set EmptyCol = New Collection
            Dim j As Integer
                For j = 1 To 6 'we're going to use the first column for this std, so 2 to 6 will be empty
                    EmptyCol.Add j
                Next j
            StdSheets.Add newStdSheet.Name, EmptyCol
            TotalEmptySpots = TotalEmptySpots + EmptyCol.Count
        Next i
        
        
        For Each std In SelectedStandards.Keys
            If TotalEmptySpots > 0 Then
                'Find the first spot to dump it
                For Each v In StdSheets 'v is the sheet name, StdSheets(v) is the EmptyCol collection
                    Set tempColl = StdSheets(v)
                    If tempColl.Count > 0 Then 'If there's a spot on sheet v,
                        Call StandardObjects(std).DumpData(Worksheets(v), False, tempColl.Item(1))
                        tempColl.Remove (1)
                        TotalEmptySpots = TotalEmptySpots - 1
                        Exit For
                    End If
                Next v
            End If
            
        Next std
                    
End Sub

Sub HarvestDumpConditionedStandard(CondStd As String, Dest As Worksheet)

    'If AllSheets Is Nothing Then Call Admin.LoadSheetLists
    Sheets("Import").Activate
    Dim Samples As Range: Set Samples = Sheets("Import").Range(Cells(2, 1), Cells(Rows.Count, 1).End(xlUp))
    Dim rng As Range
    Dim newCondStd As cProfileGenotype
    Dim newLocus As cLocusWhole
    
    For Each rng In Samples
        If rng.Value = CondStd Then
            If newCondStd Is Nothing Then 'if we haven't made the cProfileGenotype yet, make it
                Set newCondStd = New cProfileGenotype
                newCondStd.SampleName = CondStd
                newCondStd.IsStandard = True
            End If
            
            Set newLocus = CreateLocusGMID(rng.Offset(0, 1), newCondStd.IsStandard)
            newCondStd.Loci.Add newLocus.LocusName, newLocus
            
            If newLocus.SaturationFlag Then MsgBox "Peak(s) above " & SaturationMax & " RFU detected in " & newLocus.LocusName & " of " & CondStd & "." & vbNewLine & _
            "Please double-check your electropherogram.", vbExclamation, "Saturation Detected!"
            
        End If
    Next rng
    
    'Dump standard on Dest worksheet
    newCondStd.DumpData Dest, True
                     

End Sub


Sub TransferGMID(Source As Worksheet, Dest As Worksheet, NOC As Integer)
'Transfers GMID/Pre-STRmix data from a Pre-STRmix worksheet (Source) to a 1P/2P worksheet (Dest)

    Dim rngSource As Range, rngDest As Range
    
    'Transfer peaks
    Set rngSource = Source.Range("Dest_LociPreSTRmix").Offset(0, 1)
    Set rngDest = Dest.Range("Dest_" & NOC & "PLociPreSTRmix").Offset(0, 1)
    rngDest.Value = rngSource.Value
    
    'Transfer stutter
    Set rngSource = rngSource.Offset(0, 1)
    Set rngDest = rngDest.Offset(0, 1)
    rngDest.Value = rngSource.Value
    
    Dest.Range("C3").Value = Source.Range("K4").Value 'GMID Sample ID

End Sub

