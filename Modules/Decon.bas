Attribute VB_Name = "Decon"
Dim fso As New FileSystemObject
Option Explicit


Sub ImportOneDecon()

    Dim StartPath As String, DeconPath As String
    
    StartPath = Sheets("STRlite Settings").Range("STRmixResultsFolderPath").Value

    'If STRlite Settings doesn't have a valid path then just use STRlite's location
    If Not fso.FolderExists(StartPath) Then StartPath = ThisWorkbook.Path

    'Prompt for Decon folder
        With Application.FileDialog(msoFileDialogFolderPicker)
                .InitialFileName = StartPath
                If .Show = -1 Then
                    DeconPath = .SelectedItems(1)
                    Call ImportDecon(DeconPath)
                End If
        End With

End Sub

Sub ImportDecon(DeconFolder As String, Optional DestSheet As Worksheet, Optional SheetType As String = "Decon")

    Dim tempName As String, answer As Variant
    Dim VarNOC As Boolean: VarNOC = False
    Dim STRmixVersion As String
    Dim Overwrite As Boolean: Overwrite = False
    Dim fsoDecon As New FileSystemObject
    
    Dim tempCase As String, tempSample As String

    Application.ScreenUpdating = False

    'Check that the config.xml file exists and is from a Decon
        If fsoDecon.FileExists(DeconFolder & "/config.xml") Then
        
            With CreateObject("MSXML2.DOMDocument") 'don't even bother with a variable for this xml object yet
                .async = False
                .validateOnParse = False
                .Load (DeconFolder & "/config.xml")
                
                'For STRmix v2.5, the caseNumber & sampleID are only in the config file
                tempCase = .selectSingleNode("//caseNumber").Text
                tempSample = .selectSingleNode("//sampleID").Text
                
                'For all versions, the tag <mcmcSettings> indicates a decon config file
                If .selectSingleNode("//mcmcSettings") Is Nothing Then
                    MsgBox "Error validating Decon folder: " & vbNewLine & vbNewLine & DeconFolder & vbNewLine & vbNewLine & _
                        "Config.xml file is not from a deconvolution.", vbCritical + vbOKOnly, "Wrong file!"
                    Application.ScreenUpdating = True
                    Exit Sub
                End If
                
                'Check if it's a VarNOC:
                If .selectSingleNode("//maxContributors") Is Nothing Then
                    VarNOC = False
                Else:
                    VarNOC = True
                    Dim minVarNOC As Integer, maxVarNOC As Integer
                    minVarNOC = CInt(.selectSingleNode("//contributors").Text)
                    maxVarNOC = CInt(.selectSingleNode("//maxContributors").Text)
                End If
            End With
            
        Else:
            MsgBox "Error validating Decon folder: " & vbNewLine & vbNewLine & DeconFolder & vbNewLine & vbNewLine & _
                "Folder does not contain the required files.", vbCritical + vbOKOnly, "Wrong Folder?"
            Application.ScreenUpdating = True
            Exit Sub
            
        End If
    
    'Check that the results.xml file exists and is from a Decon
        If fsoDecon.FileExists(DeconFolder & "/results.xml") Then
        
            With CreateObject("MSXML2.DOMDocument")
                .async = False
                .validateOnParse = False
                .Load (DeconFolder & "/results.xml")
                
                'For all versions, the tag <analysisResult> indicates a decon result file
                If .selectSingleNode("//analysisResult") Is Nothing Then
                    MsgBox "Error validating Decon folder: " & vbNewLine & vbNewLine & DeconFolder & vbNewLine & vbNewLine & _
                        "Results.xml file is not from a deconvolution.", vbCritical + vbOKOnly, "Wrong file!"
                    Application.ScreenUpdating = True
                    Exit Sub
                End If
                
                    
                'Get STRmix Version from Results.xml file
                'Will be either "2.5" or "2.6+"
                STRmixVersion = IIf(CDbl(Left(.DocumentElement.selectSingleNode("//strmixVersion").Text, 3)) <= 2.5, "2.5", "2.6+")
                
                    
                If VarNOC = False Then
                
                    'Check if we've already imported this (non-VarNOC) decon
                    If .DocumentElement.selectSingleNode("//caseNumber") Is Nothing Then 'in STRmix v2.5 there is no caseNumber tag in the results.xml file :(
                        tempName = PubFun.FixWorksheetName(Left("(D) " & tempCase & "_" & tempSample, 31)) 'if there was no tag (v2.5) then get it from the config instead
                    Else: tempName = PubFun.FixWorksheetName(Left("(D) " & .DocumentElement.selectSingleNode("//caseNumber").Text & "_" & .DocumentElement.selectSingleNode("//sampleId").Text, 31))
                    End If
                    
                    'If there's already a sheet for that decon, ask if you want to overwrite:
                    If WorksheetExists(tempName) Then
                        answer = MsgBox("You already have a worksheet for this deconvolution:" & vbNewLine & vbNewLine & _
                            DeconFolder & vbNewLine & vbNewLine & _
                            "Do you want to overwrite it?", vbYesNo, "Overwrite Existing Deconvolution?")
                        If answer = vbYes Then Overwrite = True
                        If answer = vbNo Or answer = vbCancel Then Exit Sub
                    End If
                    
                Else: '(If VarNOC=True, stuff gets imported differently...)
                        If SheetType <> "Decon" Then
                            MsgBox DeconFolder & vbNewLine & vbNewLine & _
                            "This VarNOC sample will be saved as a regular Deconvolution Worksheet.", vbExclamation + vbOKOnly, "VarNOC"
                            SheetType = "Decon"
                        End If
                        Call ImportDeconVarNOC(DeconFolder, minVarNOC, maxVarNOC)
                        Exit Sub

                End If
                
            End With
            
        Else:
            MsgBox "Error validating Decon folder: " & vbNewLine & vbNewLine & DeconFolder & vbNewLine & vbNewLine & _
                "Folder does not contain the required files.", vbCritical + vbOKOnly, "Wrong Folder?"
            Application.ScreenUpdating = True
            Exit Sub
            
        End If
        
        
    'Create cDecon object
        Dim newDecon As cDecon
        Set newDecon = Factory.CreateDecon(DeconFolder, STRmixVersion)
        
    If DestSheet Is Nothing And SheetType = "Decon" Then
        
        'Create new Decon tab (unless overwriting)
            If Overwrite Then
                Set DestSheet = Sheets(tempName)
            Else:
                Set DestSheet = Factory.CreateDeconSheet(newDecon)
            End If
            
    End If

    'Dump everything into DestSheet
        Call newDecon.DumpData(DestSheet, SheetType)

    'Refresh Sheet lists
        If AllSheets Is Nothing Or DeconSheets Is Nothing Or SingleSheets Is Nothing Or DoubleSheets Is Nothing Then Call Admin.LoadSheetLists
        If Not AllSheets.Exists(DestSheet.Name) Then AllSheets.Add DestSheet.Name, SheetType
        
        Select Case SheetType
            Case "Decon"
                If Not DeconSheets.Exists(DestSheet.Name) Then DeconSheets.Add DestSheet.Name, "Decon"
            Case "1P"
                If Not SingleSheets.Exists(DestSheet.Name) Then SingleSheets.Add DestSheet.Name, "1P"
            Case "2P"
                If Not DoubleSheets.Exists(DestSheet.Name) Then DoubleSheets.Add DestSheet.Name, "2P"
        End Select
        
    Admin.CleanUp DestSheet
        
End Sub

Sub ImportDeconVarNOC(DeconFolder As String, minNOC As Integer, maxNOC As Integer)

    'Files in DeconFolder have already been checked at this point
    'Currently assumes two NOCs: min and max.  If STRmix ever allows a larger range in the future, I'll deal with it.
    
    Application.ScreenUpdating = False
    
    Dim tempNameMin As String, tempNameMax As String, answer As Variant
    
    'Create cDecon objects for both min & max
    Dim newDeconMin As cDecon: Set newDeconMin = Factory.CreateDecon(DeconFolder, "2.6+", minNOC)
    Dim newDeconMax As cDecon: Set newDeconMax = Factory.CreateDecon(DeconFolder, "2.6+", maxNOC)
    
    'Check if we've already imported this VarNOC decon
        tempNameMin = PubFun.FixWorksheetName(Left("(D) V" & newDeconMin.NOC & "_" & newDeconMin.CaseNum & "_" & newDeconMin.SampleID, 31))
        tempNameMax = PubFun.FixWorksheetName(Left("(D) V" & newDeconMax.NOC & "_" & newDeconMax.CaseNum & "_" & newDeconMax.SampleID, 31))
        
        'If there's already a sheet for either decon, ask if you want to overwrite
        'To keep them matched, it's all or nothing- won't let you import one non-existing VarNOC decon without overwriting the other existing one
        If PubFun.WorksheetExists(tempNameMin) Or PubFun.WorksheetExists(tempNameMax) Then
            answer = MsgBox("You already have at least one decon worksheet for this VarNOC deconvolution:" & vbNewLine & vbNewLine & _
                DeconFolder & vbNewLine & vbNewLine & _
                "STRlite will overwrite any worksheets from this VarNOC that already exist." & vbNewLine & vbNewLine & "Proceed?", vbYesNo + vbExclamation, "Overwrite Existing Deconvolution?")
                
            If answer = vbNo Or answer = vbCancel Then Exit Sub
        End If
    
    'Create new Decon tabs unless overwriting.
    Dim newDeconSheetMin As Worksheet
    Dim newDeconSheetMax As Worksheet

        'Have to check them individually in case one exists but not the other
        If PubFun.WorksheetExists(tempNameMin) Then
            Set newDeconSheetMin = ThisWorkbook.Sheets(tempNameMin)
        Else: Set newDeconSheetMin = Factory.CreateDeconSheet(newDeconMin)
        End If
        
        If PubFun.WorksheetExists(tempNameMax) Then
            Set newDeconSheetMax = ThisWorkbook.Sheets(tempNameMax)
        Else: Set newDeconSheetMax = Factory.CreateDeconSheet(newDeconMax)
        End If

    'Dump the data in each
        Call newDeconMin.DumpData(newDeconSheetMin, "Decon")
        Call newDeconMax.DumpData(newDeconSheetMax, "Decon")

    'Add the new Decon sheets to Master lists
        If AllSheets Is Nothing Or DeconSheets Is Nothing Then Call Admin.LoadSheetLists
        If Not AllSheets.Exists(newDeconSheetMin.Name) Then AllSheets.Add newDeconSheetMin.Name, "Decon"
        If Not DeconSheets.Exists(newDeconSheetMin.Name) Then DeconSheets.Add newDeconSheetMin.Name, "Decon"
        If Not AllSheets.Exists(newDeconSheetMax.Name) Then AllSheets.Add newDeconSheetMax.Name, "Decon"
        If Not DeconSheets.Exists(newDeconSheetMax.Name) Then DeconSheets.Add newDeconSheetMax.Name, "Decon"

    Application.ScreenUpdating = True

End Sub


