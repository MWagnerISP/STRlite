Attribute VB_Name = "Factory"
'Since VBA doesn't support [automatic] initialization parameters, this module contains functions
'that pass input parameters into the objects as they're created, and return those initialized objects.
'See here for explanation: https://stackoverflow.com/questions/15224113/pass-arguments-to-constructor-in-vba

'Also has the same sort of functions for creating all the new worksheets with input parameters.

'Default settings for Decon & LR live here as objects:
Public DefaultDeconSettings As cDeconSettings
Public DefaultLRSettings As cLRSettings

Option Explicit


Public Function CreateDecon(DeconFolder As String, Version As String, Optional thisNOC As Integer = 0) As cDecon

    Set CreateDecon = New cDecon
    Call CreateDecon.InitializeMe(DeconFolder, Version, thisNOC)

End Function

Public Function CreateDeconFiles(DeconFolder As String, thisNOC As Integer, Version As String) As cDeconFiles

    Set CreateDeconFiles = New cDeconFiles
    Call CreateDeconFiles.InitializeMe(DeconFolder, thisNOC, Version)
    
End Function

Public Function CreateDeconSettings(rootConfig As Object, TimeStampID As String, thisNOC As Integer) As cDeconSettings

    Set CreateDeconSettings = New cDeconSettings
    Call CreateDeconSettings.InitializeMe(rootConfig, TimeStampID, thisNOC)

End Function


Public Function CreateLR(LRFolder As String) As cLR

    Set CreateLR = New cLR
    Call CreateLR.InitializeMe(LRFolder)

End Function

Public Function CreateLRSettings(rootConfig As Object, TimeStampID As String) As cLRSettings

    Set CreateLRSettings = New cLRSettings
    Call CreateLRSettings.InitializeMe(rootConfig, TimeStampID)

End Function

Public Function CreateContributor(DeconID As String, ContributorNumber As Integer) As cContributor

    Set CreateContributor = New cContributor
    Call CreateContributor.InitializeMe(DeconID, ContributorNumber)

End Function

Public Function CreateLocusSTRmix(LocusName As String, GenotypesRange As Range) As cLocusSTRmix

    Set CreateLocusSTRmix = New cLocusSTRmix
    Call CreateLocusSTRmix.InitializeMe(LocusName, GenotypesRange)

End Function

Public Function CreateLocusGMID(MarkerCell As Range, Standard As Boolean) As cLocusWhole
'Initializes cLocusWhole from genotype table format

    Set CreateLocusGMID = New cLocusWhole
    Call CreateLocusGMID.InitializeMe(MarkerCell, Standard)

End Function

Public Function CreateLocusEV(LocusName As String) As cLocusWhole
'Initializes cLocusWhole from STRmix evidence file format

    Set CreateLocusEV = New cLocusWhole
    Call CreateLocusEV.InitializeFromSTRmixEV(LocusName)

End Function

Public Function CreateGenotype(GenotypeRange As Range, Locus As String) As cGenotypeSTRmix

    Set CreateGenotype = New cGenotypeSTRmix
    Call CreateGenotype.InitializeMe(GenotypeRange, Locus)

End Function

Public Function CreateCODIS(ProfileRange As Range, SpecifyContributors As Boolean) As cCODIS

    Set CreateCODIS = New cCODIS
    Call CreateCODIS.InitializeMe(ProfileRange, SpecifyContributors)

End Function



'****************************************************
'             Create New Worksheets
'****************************************************

Public Function CreateGMIDSheet(thisSample As cProfileGenotype) As Worksheet

    Sheets("Pre-STRmix Template").Visible = True
    Dim template As Worksheet: Set template = Sheets("Pre-STRmix Template")
    
    template.Copy After:=template
    ActiveSheet.Name = Left("(P) " & FixWorksheetName(thisSample.SampleName), 31) 'worksheet names can't be more than 31 characters long
    ActiveSheet.Tab.Color = RGB(198, 224, 180) 'green for Pre-STRmix
    ActiveSheet.Protect password:=STRlitePW, UserInterfaceOnly:=True, AllowSorting:=True, AllowFormattingCells:=True
        
    Set CreateGMIDSheet = ActiveSheet
    
    Sheets("Pre-STRmix Template").Visible = xlVeryHidden
    
End Function

Public Function CreateStandardSheet(Optional PageName As String) As Worksheet

    Sheets("Standards Template").Visible = True
    Dim template As Worksheet: Set template = Sheets("Standards Template")
    template.Name = "(Std)" 'Rename the template temporarily in case there's no PageName given
    
    template.Copy Before:=template
    ActiveSheet.Tab.Color = RGB(248, 203, 173) 'orange for Standards
    ActiveSheet.Protect password:=STRlitePW, UserInterfaceOnly:=True, AllowSorting:=True, AllowFormattingCells:=True
    
    'Rename it if applicable
    If PageName <> "" Then
        PageName = FixWorksheetName(PageName)
        If Not PubFun.WorksheetExists("(Std) " & PageName) Then ActiveSheet.Name = "(Std) " & PageName
    End If

    Set CreateStandardSheet = ActiveSheet
    template.Name = "Standards Template"
    
    Sheets("Standards Template").Visible = xlVeryHidden

End Function

Public Function CreateDeconSheet(thisDecon As cDecon) As Worksheet

    Sheets("Decon Template").Visible = True
    Dim template As Worksheet: Set template = Sheets("Decon Template")
    Dim tempName As String
    Dim btn As Variant, i As Integer
    
    template.Copy After:=template
    
    If thisDecon.VarNOC Then
        tempName = FixWorksheetName(Left("(D) V" & thisDecon.NOC & "_" & thisDecon.CaseNum & "_" & thisDecon.SampleID, 31))
    Else: tempName = FixWorksheetName(Left("(D) " & thisDecon.CaseNum & "_" & thisDecon.SampleID, 31))
    End If
    
    'Make sure tempName isn't taken; remake it until it's unique
    i = 1
    Do While PubFun.WorksheetExists(tempName)
        tempName = Left(tempName, 29) & "_" & i + 1
        i = i + 1
    Loop
        
    ActiveSheet.Name = tempName
    
    ActiveSheet.Tab.Color = RGB(155, 194, 230) 'blue
    ActiveSheet.Protect password:=STRlitePW, UserInterfaceOnly:=True, AllowSorting:=True, AllowFormattingCells:=True
    
    'Handle "Send to CODIS" buttons
    For Each btn In ActiveSheet.OLEObjects 'Turn them all off first
        If InStr(1, "CODIS", btn.Name, vbTextCompare) > 0 Then btn.Visible = False
    Next btn
    
    If thisDecon.ConditionedList.Count > 0 Then 'Turn on any conditioned contributor buttons
        For i = 1 To thisDecon.ConditionedList.Count
            ActiveSheet.OLEObjects("CondtoCODIS" & i).Visible = True
        Next i
    End If
        
    If thisDecon.NOC > thisDecon.ConditionedList.Count Then 'Turn on the rest of them up to NOC
        For i = (thisDecon.NOC - thisDecon.ConditionedList.Count) To thisDecon.NOC
            ActiveSheet.OLEObjects("ToCODIS" & i).Visible = True
        Next i
    End If
            
    Set CreateDeconSheet = ActiveSheet
    
    Sheets("Decon Template").Visible = xlVeryHidden
    
End Function

Public Function CreateLRSheet(Optional PageName As String) As Worksheet

    Sheets("LR Template").Visible = True
    Dim template As Worksheet: Set template = Sheets("LR Template")
    template.Name = "(LR)" 'Rename the template temporarily in case there's no PageName given
    
    template.Copy Before:=template
    ActiveSheet.Tab.Color = RGB(204, 204, 255) 'purple for LRs
    ActiveSheet.Protect password:=STRlitePW, UserInterfaceOnly:=True, AllowSorting:=True, AllowFormattingCells:=True
    
    'Rename it if applicable
    If PageName <> "" Then
        ActiveSheet.Range("F3:H4").Value = PageName 'dump page name as Case #
        PageName = FixWorksheetName(PageName)
        If Not PubFun.WorksheetExists("(LR) " & PageName) Then ActiveSheet.Name = "(LR) " & PageName
        
    End If

    Set CreateLRSheet = ActiveSheet
    template.Name = "LR Template"
    
    Sheets("LR Template").Visible = xlVeryHidden

End Function

Public Function CreateCODISSheet(Source As Worksheet, Optional SourceType As String = "Decon") As Worksheet

    Sheets("CODIS Template").Visible = True
    Dim template As Worksheet: Set template = Sheets("CODIS Template")
    template.Name = "(C)" 'Rename the template temporarily
    
    template.Copy Before:=template
    ActiveSheet.Tab.Color = RGB(255, 204, 255) 'pink for CODIS
    ActiveSheet.Protect password:=STRlitePW, UserInterfaceOnly:=True, AllowSorting:=True, AllowFormattingCells:=True
    
    Dim PageName As String
    
    Select Case Left(Source.Name, 3)
        Case "(D)"
            PageName = Replace(Source.Name, "(D)", "(C)", 1) 'Name the CODIS sheet after the source sheet
        Case "(1P"
            PageName = Replace(Source.Name, "(1P)", "(C)", 1)
        Case "(2P"
            PageName = Replace(Source.Name, "(2P)", "(C)", 1)
    End Select
    
    If Not PubFun.WorksheetExists(PageName) Then
        ActiveSheet.Name = PageName
        Else: ActiveSheet.Name = Left(Replace(PageName, "(C)", "(C2)"), 31) 'watch out for that 31-char limit
            'For 4-person mixtures should never need more than 2 CODIS sheets!
    End If
        
    'Automatically fill in the Case info & DeconTime from the Source
    If SourceType <> "Decon" Then
        ActiveSheet.Range("CODIS_CaseNum").Value = Source.Range("Dest_DeconResults" & SourceType).Value
        ActiveSheet.Range("CODIS_SampleID").Value = Source.Range("Dest_DeconResults" & SourceType).Offset(1, 0).Value
        ActiveSheet.Range("CODIS_DeconTime").Value = Source.Range("DeconTimestamp" & SourceType).Value
    Else:
        ActiveSheet.Range("CODIS_CaseNum").Value = Source.Range("Dest_DeconResults").Value
        ActiveSheet.Range("CODIS_SampleID").Value = Source.Range("Dest_DeconResults").Offset(1, 0).Value
        ActiveSheet.Range("CODIS_DeconTime").Value = Source.Range("DeconTimeStamp").Value
    End If
    
    Set CreateCODISSheet = ActiveSheet
    template.Name = "CODIS Template"
    
    Sheets("CODIS Template").Visible = xlVeryHidden
    
End Function

Public Function CreateComboSheet(SourceGMID As Worksheet, NOC As Integer) As Worksheet
'For now, we will only create a combo sheet from an alredy-existing Pre-STRmix sheet

    If NOC > 2 Or NOC < 1 Then Exit Function
    
    Sheets(NOC & "P Template").Visible = True
    
    Dim template As Worksheet: Set template = Sheets(NOC & "P Template")
    template.Name = "(" & NOC & "P)" 'Rename the template temporarily
    template.Copy Before:=template
    ActiveSheet.Tab.Color = RGB(255, 255, 204) 'light yellow for 1P/2P
    ActiveSheet.Protect password:=STRlitePW, UserInterfaceOnly:=True, AllowSorting:=True, AllowFormattingCells:=True
    
    Dim PageName As String
    PageName = FixWorksheetName(Replace(SourceGMID.Name, "(P)", "(" & NOC & "P)", 1)) 'Name the combo sheet after the Pre-STRmix sheet
    If Not PubFun.WorksheetExists(PageName) Then ActiveSheet.Name = PageName

    Set CreateComboSheet = ActiveSheet
    template.Name = NOC & "P Template"
    
    Sheets(NOC & "P Template").Visible = xlVeryHidden

End Function

Sub ConvertToCombo(Source As Worksheet)

    Dim NOC As Integer: NOC = Source.Range("Conts_Prop").Value
    If WorksheetExists(FixWorksheetName(Replace(Source.Name, "(P)", "(" & NOC & "P)", 1))) Then
        MsgBox "A " & NOC & "P combo worksheet already exists for this sample.", vbOKOnly, "Combo sheet exists"
        Exit Sub
    End If

    Dim newCombo As Worksheet
    Set newCombo = Factory.CreateComboSheet(Source, NOC)
    
    GMID.TransferGMID Source, newCombo, NOC
    
    Application.DisplayAlerts = False
    
    Dim answer As Variant
    answer = MsgBox("Do you want to delete the Pre-STRmix sheet?" & vbNewLine & vbNewLine & _
        "You cannot undo this.", vbYesNo + vbExclamation, "Delete Pre-STRmix Sheet?")
    
    If answer = vbYes Then Source.Delete
    
    Application.DisplayAlerts = True

End Sub

'****************************************************
'          Default Settings- Decon & LR
'****************************************************

Public Sub CreateDefaultSettings()
'Populates DefaultDeconSettings from Admin tab

    Dim Dest As Range, rng As Range

    If DefaultDeconSettings Is Nothing Then
    
        Set DefaultDeconSettings = New cDeconSettings
    
        With DefaultDeconSettings
        
            Set Dest = Sheets("STRlite Settings").Range("Settings_MCMC")
        
                .BurnAccepts = Dest.Offset(1, 0).Value
                .PostBurnAccepts = Dest.Offset(2, 0).Value
                .Chains = Dest.Offset(3, 0).Value
                .RWSD = Dest.Offset(4, 0).Value
                .PostBurnShortlist = Dest.Offset(5, 0).Value
                .MxPriors = Dest.Offset(6, 0).Value
                .AutoContinueGR = Dest.Offset(7, 0).Value
                .GRthreshold = Dest.Offset(8, 0).Value
                .ExtraAccepts = Dest.Offset(9, 0).Value
                .HRpercentAccepts = Dest.Offset(10, 0).Value
                .PopForRange = Dest.Offset(11, 0).Value
            
            Set Dest = Sheets("STRlite Settings").Range("Settings_Kit")
            
                .KitName = Dest.Offset(1, 0).Value
                .AutosomalLoci = Dest.Offset(2, 0).Value
                .Saturation = Dest.Offset(3, 0).Value
                .DegradStart = Dest.Offset(4, 0).Value
                .DegradMax = Dest.Offset(5, 0).Value
                .DropInCap = Dest.Offset(6, 0).Value
                .DropInFreq = Dest.Offset(7, 0).Value
                .DropInGamma = Dest.Offset(8, 0).Value
                .AlleleVariance = Dest.Offset(9, 0).Value
                .MinVarianceFactor = Dest.Offset(10, 0).Value
                .LocusAmpVariance = Dest.Offset(11, 0).Value
                .VarMinParameter = Dest.Offset(12, 0).Value
            
            
            'Up to 4 Types of Stutter
                Set Dest = Sheets("STRlite Settings").Range("Settings_Stutter")
                Dim newStutter As cStutterModel
                
                For Each rng In Dest
                    If rng.Value <> "" And Not .StutterSettings.Exists(rng.Value) Then
                        Set newStutter = New cStutterModel
                        
                        With newStutter
                            .StutterName = rng.Value
                            .StutterMax = rng.Offset(0, 1).Value
                            .StutterVariance = rng.Offset(0, 2).Value
                        End With
                        
                        .StutterSettings.Add newStutter.StutterName, newStutter
                        
                    End If
                Next rng
                
        End With
    
    End If

End Sub


Public Sub CreateDefaultLRSettings()

    Dim Dest As Range

    If DefaultLRSettings Is Nothing Then
    
        Set DefaultLRSettings = New cLRSettings
    
        Set Dest = Sheets("STRlite Settings").Range("Settings_LR")
        
        With DefaultLRSettings
        
            .AssignSubSourceLR = Dest.Offset(1, 0).Value
            .CalculateHPD = Dest.Offset(2, 0).Value
            .HPDiterations = Dest.Offset(3, 0).Value
            .MCMCuncertainty = Dest.Offset(4, 0).Value
            .AlleleFreqUncertainty = Dest.Offset(5, 0).Value
            .HPDquantile = Dest.Offset(6, 0).Value
            .HPDsides = Dest.Offset(7, 0).Value
            
        End With
        
    End If

End Sub
