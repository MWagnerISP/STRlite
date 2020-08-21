VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pTargetSheet As String
Private pCaseNum As String
Private pSampleID As String
Private pNOC As String
Private pEvidenceFileName As String

Private pVarNOC As Boolean
Private pHpNOC As Integer
Private pHdNOC As Integer

Private pFolderPath_LR As String 'Full path of LR folder
Private pFolderName_LR As String 'Full name of LR folder
Private pFilePath_LRConfigXML As String 'Filepath of Config.xml file
Private pFilePath_LRResultsXML As String 'Filepath of Results.xml file
Private pFilePath_LRResultsTXT As String 'Filepath of Results.txt file

Private pTimeStamp_DeconXML As String
Private pTimeStamp_LR As String

Private pLRSettings As New cLRSettings
Private pSettings_Check As Boolean

Private pStratifiedTotal As Double
Private pStratifiedUnified As Double

Private pTotalHPD As New Scripting.Dictionary 'key = population, item = HPD LR value
Private pTotalPointLR As New Scripting.Dictionary 'key = population item = point LR value

Private pIgnoredLoci As New Scripting.Dictionary 'key = locus
'Private pPopulations As New Scripting.Dictionary 'key = population
Private pHpKnowns As New Scripting.Dictionary 'key = reference name
Private pHdKnowns As New Scripting.Dictionary 'key = reference name

Dim LRResultsXML As IXMLDOMDocument 'actual xml file loaded as object
Dim LRConfigXML As IXMLDOMDocument  'actual xml file loaded as object
Dim rootResults As IXMLDOMElement  'root element of LRResultsXML
Dim rootConfig As IXMLDOMElement  'IXMLDOMElement root element of LRConfigXML

Option Explicit


Public Sub InitializeMe(LRFolder As String)

    'Grab file/folder paths
    Call Me.LRFiles(LRFolder)
    
    'Import XML files/load into memory
    'Also detect VarNOC here
    Call Me.ImportData
    
    If rootResults Is Nothing Or rootConfig Is Nothing Then Exit Sub
    
    'Grab timestamps & format them
        Dim tempStamp As String
        'v2.5 uses "runDate" not "interpretationRunDate"
        If rootResults.selectSingleNode("//interpretationRunDate") Is Nothing Then
            'in v2.5, the decon timestamp in the LR file has to come from the decon folder name in "previousAnalysis".
            'Except "previousAnalysis" doesn't exist for combined decons/LRs, and "runDate" isn't unique to decons, so...
            If rootConfig.selectSingleNode("//previousAnalysis") Is Nothing Then 'if previousAnalysis doesn't exist, it's a combined decon/LR
                tempStamp = rootResults.selectSingleNode("//runDate").Text
                Me.TimeStamp_DeconXML = Replace((Left(tempStamp, 10) & "-" & Mid(tempStamp, 12, 8)), ":", "-") 'if it's a combined decon/LR, runDate is same for both
            Else:
                Me.TimeStamp_DeconXML = Right(rootConfig.selectSingleNode("//previousAnalysis").Text, 19)
            End If
            
        Else:
            tempStamp = rootResults.selectSingleNode("//interpretationRunDate").Text 'interpretationRunDate for v2.6+ Decon
            Me.TimeStamp_DeconXML = Replace((Left(tempStamp, 10) & "-" & Mid(tempStamp, 12, 8)), ":", "-")
        End If
        
        tempStamp = rootResults.selectSingleNode("//runDate").Text 'runDate for LR
        Me.TimeStamp_LR = Replace((Left(tempStamp, 10) & "-" & Mid(tempStamp, 12, 8)), ":", "-")
    
    'Create cLRSettings object and harvest settings from DeconConfigXML/rootConfig
    Set Me.LRSettings = Factory.CreateLRSettings(rootConfig, Me.TimeStamp_DeconXML)
    
    'Check the LRSettings against defaults and save as Me.Settings_Check
    Me.Settings_Check = Me.LRSettings.CheckAgainstDefault
    
    'Harvest results from LRResultsXML
    Call Me.HarvestLR

End Sub

Sub LRFiles(LRFolder As String)

    Dim fsoLR As FileSystemObject: Set fsoLR = CreateObject("Scripting.FileSystemObject")
    
    With Me
        .FolderPath_LR = LRFolder 'Full path of LR folder containing Results.txt file
        .FolderName_LR = fsoLR.GetFileName(LRFolder) 'Full name of LR folder
        .FilePath_LRConfigXML = LRFolder & "/config.xml" 'Filepath of Config.xml file
        .FilePath_LRResultsXML = LRFolder & "/results.xml" 'Filepath of Results.xml file
    End With
    
End Sub

Sub ImportData()
    
    Application.DisplayAlerts = False

    'Import LR Results xml
    Set LRResultsXML = CreateObject("MSXML2.DOMDocument")
        With LRResultsXML
            .async = False
            .validateOnParse = False
            .Load (Me.FilePath_LRResultsXML)
        End With
    Set rootResults = LRResultsXML.DocumentElement
    
    'Import LR Config xml
    Set LRConfigXML = CreateObject("MSXML2.DOMDocument")
        With LRConfigXML
            .async = False
            .validateOnParse = False
            .Load (Me.FilePath_LRConfigXML)
        End With
    Set rootConfig = LRConfigXML.DocumentElement
    
    
    
    Me.CaseNum = rootConfig.selectSingleNode("//caseNumber").Text
    
    'Detect VarNOC
    If rootConfig.selectSingleNode(".//varNocPriorsMethod") Is Nothing Then
        Me.VarNOC = False
        Else: Me.VarNOC = True
    End If
    
    Me.SampleID = rootConfig.selectSingleNode("//sampleID").Text 'STRmix changes the SampleID between config and results!!! :-(
    
    'Import Results.txt file
        'The identities/#s of assumed contributors is only in the results.txt file. Please add to xml file, NicheVision!
        'The Import tab is used temporarily to store data from results.txt, so it is cleared & reused with each LR.
    Me.FilePath_LRResultsTXT = Me.FolderPath_LR & "/" & Dir(Me.FolderPath_LR & "/*Results.txt") 'path of Results.txt file

    Dim wkbTemp As Workbook
    Dim CurrentFile As String: CurrentFile = ThisWorkbook.Name
    
    Workbooks(CurrentFile).Sheets("Import").Cells.ClearContents
    
    Application.ScreenUpdating = False
    
'    Set wkbTemp = Workbooks.Open(Me.FilePath_LRResultsTXT)
'    wkbTemp.Sheets(1).Range("A:A").Copy
'    Workbooks(CurrentFile).Activate
    
    
    With Workbooks.Open(Me.FilePath_LRResultsTXT)
        .Sheets(1).UsedRange.Copy destination:=Workbooks(CurrentFile).Sheets("Import").Range("A1")
        .Close False
    End With
    
    
    
'    Sheets("Import").Activate
'    Sheets("Import").Range("A1").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'    :=False, Transpose:=False
    
    Application.CutCopyMode = False
    'wkbTemp.Close
      
    Application.DisplayAlerts = True
    
End Sub

Sub HarvestLR()

    Dim v As Variant, j As Integer, vNode As IXMLDOMNode
    Dim vListPopLRs As IXMLDOMNodeList, popname As String, HPD As String, pointLR As String
    Dim stratTotalNode As IXMLDOMNode, strStratUnifiedNode As String
    
    
    'Get Ignored Loci (from config)
        Dim vListIgnored As IXMLDOMNodeList 'IXMLDOMNodeList returned from .SelectNodes method
        
        'The list of ignored loci is in a slightly different hierarchy depending on Decon/LR combined vs. LrFromPrevious
        If rootConfig.BaseName = "lrFromPrevious" Then Set vListIgnored = rootConfig.SelectNodes("//ignoredLoci/kit/locus")
        If rootConfig.BaseName = "analysis" Then Set vListIgnored = rootConfig.SelectNodes("//ignoredLoci/locus")
        
        For Each vNode In vListIgnored
            'fill dictionary (.IgnoredLoci) from IXMLDOMNodeList, but only if the locus doesn't have a "Y" (Y locus)
            If InStr(1, vNode.Text, "Y", vbBinaryCompare) = 0 Then Me.IgnoredLoci.Add vNode.Text, Nothing
        Next vNode
        
    
    'Handle VarNOC
    If VarNOC Then
        'Detect type of VarNOC
        Dim TypeVarNOC As String: TypeVarNOC = rootConfig.selectSingleNode("//varNocPriorsMethod").Text
        
        Select Case TypeVarNOC
        
            Case "MLE"
                Me.HpNOC = CInt(rootResults.selectSingleNode("//numberOfContributorsUnderHp").Text)
                Me.HdNOC = CInt(rootResults.selectSingleNode("//numberOfContributorsUnderHd").Text)
                Me.NOC = IIf(Me.HpNOC = Me.HdNOC, Me.HpNOC, Me.HpNOC & "/" & Me.HdNOC)
                   
            Case "Stratified"
                'Stratified doesn't "pick" a particular NOC at the end
                Dim vListNOCs As IXMLDOMNodeList
                Set vListNOCs = rootResults.SelectNodes("//contributorStratifiedLr/contributorStratifiedLr")
                Me.NOC = vListNOCs.Item(0).Attributes(0).Text & "/" & vListNOCs.Item(1).Attributes(0).Text
                
            Case "User Selected"
                Me.HpNOC = CInt(rootResults.selectSingleNode("//numberOfContributorsUnderHp").Text)
                Me.HdNOC = CInt(rootResults.selectSingleNode("//numberOfContributorsUnderHd").Text)
                Me.NOC = IIf(Me.HpNOC = Me.HdNOC, Me.HpNOC, Me.HpNOC & "/" & Me.HdNOC)
            
        End Select
        
        Set vListPopLRs = rootResults.SelectNodes("//contributorStratifiedLrResult/lrSummary/populations/population/relationLR/relation[@type='TOTAL']")
        Set stratTotalNode = rootResults.selectSingleNode("//contributorStratifiedLrResult/lrSummary/stratified/relationLR/relation[@type='TOTAL']")
        strStratUnifiedNode = "//contributorStratifiedLrResult/lrSummary/stratified/relationLR/relation[@type='HD_UNIFIED']"
    
    Else:
    'Non-VarNOC
        Set vListPopLRs = rootResults.SelectNodes("//population/relationLR/relation[@type='TOTAL']")
        Set stratTotalNode = rootResults.selectSingleNode("//stratified/relationLR/relation[@type='TOTAL']") 'there will always be a "TOTAL" node
        strStratUnifiedNode = "//stratified/relationLR/relation[@type='HD_UNIFIED']"
        
        'v2.5 has the NOC in an attribute of "lrFromPrevResult" instead of its own node. Not sure if the parent node changes when Decon/LR runs together.
        If rootResults.selectSingleNode("//numberOfContributors/contributors") Is Nothing Then
            Me.NOC = rootResults.selectSingleNode("//lrSummary").ParentNode.Attributes(0).Text 'select the parent node of lrSummary and then take that node's attribute's text
            Else: Me.NOC = rootResults.selectSingleNode("//numberOfContributors/contributors").Text
        End If
        
    End If
        
        
    'HPD/pointLRs per population
        For Each vNode In vListPopLRs
            popname = vNode.ParentNode.ParentNode.Attributes(0).Text
            
            If Me.LRSettings.CalculateHPD Then
                HPD = CDbl(vNode.selectSingleNode("hpdValue").Text)
                If Not Me.TotalHPD.Exists(popname) Then Me.TotalHPD.Add popname, HPD
            End If
            
            pointLR = CDbl(vNode.selectSingleNode("lrValue").Text)
            If Not Me.TotalPointLR.Exists(popname) Then Me.TotalPointLR.Add popname, pointLR
        Next vNode
        
        
    'Stratified total & stratified-unified
        'If subsource LR is turned off, there will be no Unified calc, so just get the Total
        If rootResults.selectSingleNode(strStratUnifiedNode) Is Nothing Then
            If Me.LRSettings.CalculateHPD Then
                Me.StratifiedTotal = CDbl(stratTotalNode.selectSingleNode("hpdValue").Text)
            Else:
                Me.StratifiedTotal = CDbl(stratTotalNode.selectSingleNode("lrValue").Text)
            End If
        Else:
            Dim stratUnifiedNode As IXMLDOMNode: Set stratUnifiedNode = rootResults.selectSingleNode(strStratUnifiedNode)
            If Me.LRSettings.CalculateHPD Then
                Me.StratifiedTotal = CDbl(stratTotalNode.selectSingleNode("hpdValue").Text)
                Me.StratifiedUnified = CDbl(stratUnifiedNode.selectSingleNode("hpdValue").Text)
            Else:
                Me.StratifiedTotal = CDbl(stratTotalNode.selectSingleNode("lrValue").Text)
                Me.StratifiedUnified = CDbl(stratUnifiedNode.selectSingleNode("lrValue").Text)
            End If
        End If
        
        
        
    'Hp contributors:
    
    Dim rngHp As Range, HpCount As Integer, HpProps() As String
    Set rngHp = PubFun.FindCell(Sheets("Import").Range("A:A"), "Number of known contributors under Hp:", Range("A1"))
        
    HpCount = Val(Right(rngHp.Value, 1)) '# of contributors under Hp

    If rngHp.Offset(1, 0).Value <> "" Then
        HpProps() = Split(rngHp.Offset(1, 0).Value, ", ")
        'get rid of unnecessary text (including case #, if same as decon case #)
        For j = LBound(HpProps) To UBound(HpProps)
            If InStr(1, HpProps(j), ".csv") Then HpProps(j) = Replace(HpProps(j), ".csv", "")
            If InStr(1, HpProps(j), "_REF") Then HpProps(j) = Replace(HpProps(j), "_REF", "")
            If InStr(1, HpProps(j), Me.CaseNum & "_") Then HpProps(j) = Replace(HpProps(j), Me.CaseNum & "_", "")
            If InStr(1, HpProps(j), Me.CaseNum & " ") Then HpProps(j) = Replace(HpProps(j), Me.CaseNum & " ", "")
            If Not Me.HpKnowns.Exists(j) And HpProps(j) <> "" Then Me.HpKnowns.Add HpProps(j), Nothing
        Next j
    End If
    
    
    'Hd Contributors:
    
    Dim rngHd As Range, HdCount As Integer, HdProps() As String
    Set rngHd = PubFun.FindCell(Sheets("Import").Range("A:A"), "Number of known contributors under Hd:", Range("A1"))
        
    HdCount = Val(Right(rngHd.Value, 1)) '# of contributors under Hd

    If rngHd.Offset(1, 0).Value <> "" Then
        HdProps() = Split(rngHd.Offset(1, 0).Value, ", ")
        'get rid of unnecessary text (including case #, if same as decon case #)
        For j = LBound(HdProps) To UBound(HdProps)
            If InStr(1, HdProps(j), ".csv") Then HdProps(j) = Replace(HdProps(j), ".csv", "")
            If InStr(1, HdProps(j), "_REF") Then HdProps(j) = Replace(HdProps(j), "_REF", "")
            If InStr(1, HdProps(j), Me.CaseNum & "_") Then HdProps(j) = Replace(HdProps(j), Me.CaseNum & "_", "")
            If InStr(1, HdProps(j), Me.CaseNum & " ") Then HdProps(j) = Replace(HdProps(j), Me.CaseNum & " ", "")
            If Not Me.HdKnowns.Exists(j) And HdProps(j) <> "" Then Me.HdKnowns.Add HdProps(j), Nothing
        Next j
    End If
    
    
    'Evidence File Name
    Dim rngEvidence As Range: Set rngEvidence = PubFun.FindCell(Sheets("Import").Range("A:A"), "Evidence sample", Range("A1"))
    Me.EvidenceFileName = Replace(Replace(rngEvidence.Value, "Evidence sample ", "", 1), ".csv", "")
    

End Sub

Sub DumpLR1P(DestSheet As Worksheet)

    Dim DestSampleID As String: DestSampleID = "I21"
    Dim DestFileMatch As String: DestFileMatch = "I22"
    Dim DestSettingCheck As String: DestSettingCheck = "I23"
    Dim DestIgnoredLoci As String: DestIgnoredLoci = "I24"

    If CInt(Me.NOC) <> 1 Or Me.VarNOC = True Then Exit Sub
    
    'Check for file match
    If Me.TimeStamp_DeconXML = DestSheet.Range("S2").Value Then
        DestSheet.Range(DestFileMatch).Value = "File match verified."
    Else:
        MsgBox "This LR appears to be from a different decon." & vbNewLine & vbNewLine & "The decon timestamp does not match.", vbOKOnly + vbExclamation, "Wrong LR?"
        Exit Sub
    End If
    
    'Ignored loci:
        Dim strIgnored As String
        Select Case Me.IgnoredLoci.Count
            Case 0
            strIgnored = "" 'Just keep it blank if nothing was omitted. We don't have space to waste.
            
            Case Is > 4
            strIgnored = "Autosomal loci omitted. Check report."
            
            Case Else
            strIgnored = "Omitted " & Join(IgnoredLoci.Keys(), ", ")
        End Select
        
    With DestSheet
        .Range(DestSampleID).Value = Me.SampleID
        .Range(DestSettingCheck).Value = IIf(Me.Settings_Check, "Default settings verified.", "Check settings.")
        .Range(DestIgnoredLoci).Value = strIgnored
        .Range("I26").Value = Me.HpKnowns.Keys()(0) 'single source LR means one HpKnown
        .Range("I27").Value = "U" 'single source LR means Hd = "U"
        
        Dim LRcorner As Range: Set LRcorner = .Range("AA3") 'LRcorner is the first HPD per population spot
        Dim i As Integer: i = 0
        Dim v As Variant
    
        'HPDs per population
        i = 0
        For Each v In Me.TotalHPD.Items
            LRcorner.Offset(0, i).Value = CDbl(v)
            i = i + 1
        Next v
        
        'Point LRs per population
        i = 0
        For Each v In Me.TotalPointLR.Items
            LRcorner.Offset(4, i).Value = CDbl(v)
            i = i + 1
        Next v
        
        'Stratified
        LRcorner.Offset(5, 2).Value = Me.StratifiedTotal
        LRcorner.Offset(6, 2).Value = Me.StratifiedUnified
        
    End With

End Sub

Sub DumpLR2P(DestSheet As Worksheet)

    Dim DestSampleID As String: DestSampleID = "I23"
    Dim DestFileMatch As String: DestFileMatch = "I24"
    Dim DestSettingCheck As String: DestSettingCheck = "I25"
    Dim DestIgnoredLoci As String: DestIgnoredLoci = "I26"
    
    If CInt(Me.NOC) <> 2 Or Me.VarNOC = True Then Exit Sub
    
    'Check for file match
    If Me.TimeStamp_DeconXML = DestSheet.Range("S2").Value Then
        DestSheet.Range(DestFileMatch).Value = "File match verified."
    Else:
        MsgBox "This LR appears to be from a different decon." & vbNewLine & vbNewLine & "The decon timestamp does not match.", vbOKOnly + vbExclamation, "Wrong LR?"
        Exit Sub
    End If
    
    'Ignored loci:
        Dim strIgnored As String
        Select Case Me.IgnoredLoci.Count
            Case 0
            strIgnored = "" 'Just keep it blank if nothing was omitted. We don't have space to waste.
            
            Case Is > 4
            strIgnored = "Autosomal loci omitted. Check report."
            
            Case Else
            strIgnored = "Omitted " & Join(IgnoredLoci.Keys(), ", ")
        End Select

    With DestSheet
        .Range(DestSampleID).Value = Me.SampleID
        .Range(DestSettingCheck).Value = IIf(Me.Settings_Check, "Default settings verified.", "Check settings.")
        .Range(DestIgnoredLoci).Value = strIgnored
        
        Select Case Me.HpKnowns.Count 'will be either 1 or 2
            Case 1
                .Range("I28").Value = Me.HpKnowns.Keys()(0)
                .Range("K28").Value = "U"
            Case 2
                .Range("I28").Value = Me.HpKnowns.Keys()(0)
                .Range("K28").Value = Me.HpKnowns.Keys()(1)
        End Select
        
        Select Case Me.HdKnowns.Count 'will be either 0 or 1
            Case 0
                .Range("I29").Value = "U"
                .Range("K29").Value = "U"
            Case 1
                .Range("I29").Value = Me.HdKnowns.Keys()(0)
                .Range("K29").Value = "U"
        End Select


        Dim LRcorner As Range: Set LRcorner = .Range("AA3") 'LRcorner is the first HPD per population spot
        Dim i As Integer: i = 0
        Dim v As Variant
    
        'HPDs per population
        i = 0
        For Each v In Me.TotalHPD.Items
            LRcorner.Offset(0, i).Value = CDbl(v)
            i = i + 1
        Next v
        
        'Point LRs per population
        i = 0
        For Each v In Me.TotalPointLR.Items
            LRcorner.Offset(4, i).Value = CDbl(v)
            i = i + 1
        Next v
        
        'Stratified
        LRcorner.Offset(5, 2).Value = Me.StratifiedTotal
        LRcorner.Offset(6, 2).Value = Me.StratifiedUnified
        
    End With
    
End Sub


Sub DumpLR(LRSheet As Worksheet, rngCorner As Range)
'rngCorner and LRSheet are inherited all the way from the Import LR button on the LR Template page

    Dim v As Variant, i As Integer, rng As Range, LRcorner As Range
    
    Set rngCorner = LRSheet.Range(rngCorner.Address) 'Make sure we're targeting rngCorner on the right worksheet
    Set LRcorner = rngCorner.Offset(3, 14) 'LRcorner is the first HPD per population spot
    
    'Clear LRcorner area
    LRcorner.Resize(1, 4).ClearContents
    LRcorner.Offset(4, 0).Resize(1, 4).ClearContents
    LRcorner.Offset(5, 2).ClearContents
    LRcorner.Offset(6, 2).ClearContents

    'Basics:
        rngCorner.Offset(0, 2).Value = Me.SampleID
        rngCorner.Offset(6, 2).Value = Me.EvidenceFileName
        rngCorner.Offset(7, 2).Value = Me.TimeStamp_DeconXML
        rngCorner.Offset(8, 1).Value = IIf(Me.Settings_Check, "Default settings verified.", "Check settings.")
        rngCorner.Offset(8, 4).Value = IIf(Me.VarNOC, Me.NOC & " (Var)", Me.NOC)
    
    'Ignored loci:
        Dim strIgnored As String
        Select Case Me.IgnoredLoci.Count
            
            Case 0
            strIgnored = "No autosomal loci omitted."
            
            Case Is > 4
            strIgnored = "Autosomal loci omitted. Check report."
            
            Case Else
            strIgnored = "Omitted " & Join(IgnoredLoci.Keys(), ", ")
            
        End Select
        
        rngCorner.Offset(9, 1).Value = strIgnored
    
    'LR values
    
        'HPDs per population
        i = 0
        For Each v In Me.TotalHPD.Items
            LRcorner.Offset(0, i).Value = CDbl(v)
            i = i + 1
        Next v
        
        'Point LRs per population
        i = 0
        For Each v In Me.TotalPointLR.Items
            LRcorner.Offset(4, i).Value = CDbl(v)
            i = i + 1
        Next v
        
        'Stratified
        LRcorner.Offset(5, 2).Value = Me.StratifiedTotal
        LRcorner.Offset(6, 2).Value = Me.StratifiedUnified
        
        
    'Propositions (H1/H2)
        If Me.VarNOC And Me.HdNOC <> Me.HpNOC Then
            Call Me.WeirdVarNOC(rngCorner) 'If H1 & H2 have different NOC!
        Else:
            Call Me.ContributorSection(rngCorner)
        End If

End Sub

Sub ContributorSection(rng As Range)
'rng is the main LR location (rngCorner, upper left cell near Sample ID)

    Dim v As Variant, i As Integer, col As Integer
    Dim Target As Range: Set Target = rng.Offset(3, 0)
    'Target = the cell to the left of the main contributor area (contains "H1")
    
    'Unmerge everything & background = white
    With Target.Offset(0, 1).Resize(2, 4)
        .UnMerge
        .Interior.Color = RGB(255, 255, 255)
        .Font.Name = "Perpetua"
        .Font.Size = 12
    End With
    
    
    Select Case Me.NOC
    
        Case 1

            Target.Offset(0, 1).Value = Me.HpKnowns.Keys()(0) 'always at least 1 HpKnown
            Target.Offset(0, 1).Resize(1, 4).Merge

            Target.Offset(1, 1).Value = "U" 'shouldn't be any HdKnowns when NOC=1
            Target.Offset(1, 1).Resize(1, 4).Merge
            
        Case 2
        
            For col = 1 To 3 Step 2
                Target.Offset(0, col).Value = "U"
                Target.Offset(0, col).Resize(1, 2).Merge
                Target.Offset(1, col).Value = "U"
                Target.Offset(1, col).Resize(1, 2).Merge
            Next col
            
            For col = HpKnowns.Count To 1 Step -1
                Target.Offset(0, (2 * col) - 1).Value = HpKnowns.Keys()(col - 1)
            Next col
            
            For col = HdKnowns.Count To 1 Step -1
                Target.Offset(1, (2 * col) - 1).Value = HdKnowns.Keys()(col - 1)
            Next col
            
            
        Case 3
        
            For col = 1 To 3
                Target.Offset(0, col).Value = "U"
                Target.Offset(1, col).Value = "U"
            Next col
            
            For col = 1 To HpKnowns.Count
                Target.Offset(0, col).Value = HpKnowns.Keys()(col - 1)
            Next col
            
            For col = 1 To HdKnowns.Count
                Target.Offset(1, col).Value = HdKnowns.Keys()(col - 1)
            Next col
                
            'Color & merge last (4th) column only
            With Target.Offset(0, 4).Resize(2, 1)
                .Interior.Color = Target.Offset(-1, 0).Interior.Color
                .Merge
            End With
        
        Case 4
        
            For col = 1 To 4
                Target.Offset(0, col).Value = "U"
                Target.Offset(1, col).Value = "U"
            Next col
            
            For col = 1 To HpKnowns.Count
                Target.Offset(0, col).Value = HpKnowns.Keys()(col - 1)
            Next col
            
            For col = 1 To HdKnowns.Count
                Target.Offset(1, col).Value = HdKnowns.Keys()(col - 1)
            Next col
        
    End Select
    

End Sub


Sub WeirdVarNOC(rng As Range)
'This handles NOC for H1 & H2 separately

    Dim v As Variant, i As Integer, col As Integer
    Dim Target As Range: Set Target = rng.Offset(3, 0)
    'Target = the cell to the left of the main contributor area (contains "H1")
    
    'Unmerge everything & background = white
    With Target.Offset(0, 1).Resize(2, 4)
        .UnMerge
        .Interior.Color = RGB(255, 255, 255)
        .Font.Name = "Perpetua"
        .Font.Size = 12
    End With
    
    Select Case Me.HpNOC
        Case 1

            Target.Offset(0, 1).Value = Me.HpKnowns.Keys()(0) 'always at least 1 HpKnown
            Target.Offset(0, 1).Resize(1, 4).Merge

        Case 2
        
            For col = 1 To 3 Step 2
                Target.Offset(0, col).Value = "U"
                Target.Offset(0, col).Resize(1, 2).Merge
            Next col
            
            For col = HpKnowns.Count To 1 Step -1
                Target.Offset(0, (2 * col) - 1).Value = HpKnowns.Keys()(col - 1)
            Next col
            
        Case 3
        
            For col = 1 To 3
                Target.Offset(0, col).Value = "U"
            Next col
            
            For col = 1 To HpKnowns.Count
                Target.Offset(0, col).Value = HpKnowns.Keys()(col - 1)
            Next col

        Case 4
        
            For col = 1 To 4
                Target.Offset(0, col).Value = "U"
            Next col
            
            For col = 1 To HpKnowns.Count
                Target.Offset(0, col).Value = HpKnowns.Keys()(col - 1)
            Next col
        
    End Select
    
    Select Case Me.HdNOC
        Case 1
            Target.Offset(1, 1).Value = "U" 'shouldn't be any HdKnowns when NOC=1
            Target.Offset(1, 1).Resize(1, 4).Merge
    
        Case 2
            For col = 1 To 3 Step 2
                Target.Offset(1, col).Value = "U"
                Target.Offset(1, col).Resize(1, 2).Merge
            Next col
                        
            For col = HdKnowns.Count To 1 Step -1
                Target.Offset(1, (2 * col) - 1).Value = HdKnowns.Keys()(col - 1)
            Next col
        
        Case 3
            For col = 1 To 3
                Target.Offset(1, col).Value = "U"
            Next col
                    
            For col = 1 To HdKnowns.Count
                Target.Offset(1, col).Value = HdKnowns.Keys()(col - 1)
            Next col
        
        Case 4
            For col = 1 To 4
                Target.Offset(1, col).Value = "U"
            Next col
                    
            For col = 1 To HdKnowns.Count
                Target.Offset(1, col).Value = HdKnowns.Keys()(col - 1)
            Next col
            
    End Select

End Sub



Public Property Get TargetSheet() As String
    TargetSheet = pTargetSheet
End Property
Public Property Let TargetSheet(v As String)
    pTargetSheet = v
End Property

Public Property Get CaseNum() As String
    CaseNum = pCaseNum
End Property
Public Property Let CaseNum(v As String)
    pCaseNum = v
End Property

Public Property Get SampleID() As String
    SampleID = pSampleID
End Property
Public Property Let SampleID(v As String)
    pSampleID = v
End Property

Public Property Get NOC() As String
    NOC = pNOC
End Property
Public Property Let NOC(v As String)
    pNOC = v
End Property

Public Property Get EvidenceFileName() As String
    EvidenceFileName = pEvidenceFileName
End Property
Public Property Let EvidenceFileName(v As String)
    pEvidenceFileName = v
End Property

Public Property Get VarNOC() As Boolean
    VarNOC = pVarNOC
End Property
Public Property Let VarNOC(v As Boolean)
    pVarNOC = v
End Property

Public Property Get HpNOC() As Integer
    HpNOC = pHpNOC
End Property
Public Property Let HpNOC(v As Integer)
    pHpNOC = v
End Property

Public Property Get HdNOC() As Integer
    HdNOC = pHdNOC
End Property
Public Property Let HdNOC(v As Integer)
    pHdNOC = v
End Property


Public Property Get TimeStamp_DeconXML() As String
    TimeStamp_DeconXML = pTimeStamp_DeconXML
End Property
Public Property Let TimeStamp_DeconXML(v As String)
    pTimeStamp_DeconXML = v
End Property

Public Property Get TimeStamp_LR() As String
    TimeStamp_LR = pTimeStamp_LR
End Property
Public Property Let TimeStamp_LR(v As String)
    pTimeStamp_LR = v
End Property




Public Property Get FolderName_LR() As String
    FolderName_LR = pFolderName_LR
End Property
Public Property Let FolderName_LR(v As String)
    pFolderName_LR = v
End Property

Public Property Get FolderPath_LR() As String
    FolderPath_LR = pFolderPath_LR
End Property
Public Property Let FolderPath_LR(v As String)
    pFolderPath_LR = v
End Property

Public Property Get FilePath_LRConfigXML() As String
    FilePath_LRConfigXML = pFilePath_LRConfigXML
End Property
Public Property Let FilePath_LRConfigXML(v As String)
    pFilePath_LRConfigXML = v
End Property

Public Property Get FilePath_LRResultsXML() As String
    FilePath_LRResultsXML = pFilePath_LRResultsXML
End Property
Public Property Let FilePath_LRResultsXML(v As String)
    pFilePath_LRResultsXML = v
End Property

Public Property Get FilePath_LRResultsTXT() As String
    FilePath_LRResultsTXT = pFilePath_LRResultsTXT
End Property
Public Property Let FilePath_LRResultsTXT(v As String)
    pFilePath_LRResultsTXT = v
End Property


'Setting Checks

Public Property Get Settings_Check() As Boolean
    Settings_Check = pSettings_Check
End Property
Public Property Let Settings_Check(v As Boolean)
    pSettings_Check = v
End Property

Public Property Get LRSettings() As cLRSettings
    Set LRSettings = pLRSettings
End Property
Public Property Set LRSettings(v As cLRSettings)
    Set pLRSettings = v
End Property


Public Property Get StratifiedTotal() As Double
    StratifiedTotal = pStratifiedTotal
End Property
Public Property Let StratifiedTotal(v As Double)
    pStratifiedTotal = v
End Property

Public Property Get StratifiedUnified() As Double
    StratifiedUnified = pStratifiedUnified
End Property
Public Property Let StratifiedUnified(v As Double)
    pStratifiedUnified = v
End Property

Public Property Get TotalHPD() As Scripting.Dictionary
    Set TotalHPD = pTotalHPD
End Property
Public Property Let TotalHPD(v As Scripting.Dictionary)
    Set pTotalHPD = v
End Property

Public Property Get TotalPointLR() As Scripting.Dictionary
    Set TotalPointLR = pTotalPointLR
End Property
Public Property Let TotalPointLR(v As Scripting.Dictionary)
    Set pTotalPointLR = v
End Property


Public Property Get IgnoredLoci() As Scripting.Dictionary
    Set IgnoredLoci = pIgnoredLoci
End Property
Public Property Let IgnoredLoci(v As Scripting.Dictionary)
    Set pIgnoredLoci = v
End Property

'Public Property Get Populations() As Scripting.Dictionary
'    Set Populations = pPopulations
'End Property
'Public Property Let Populations(v As Scripting.Dictionary)
'    Set pPopulations = v
'End Property

Public Property Get HpKnowns() As Scripting.Dictionary
    Set HpKnowns = pHpKnowns
End Property
Public Property Let HpKnowns(v As Scripting.Dictionary)
    Set pHpKnowns = v
End Property

Public Property Get HdKnowns() As Scripting.Dictionary
    Set HdKnowns = pHdKnowns
End Property
Public Property Let HdKnowns(v As Scripting.Dictionary)
    Set pHdKnowns = v
End Property


