VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDecon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Represents a single Deconvolution run through STRmix. Includes both decon-only and decon/LR combos.
'This is a hefty class. Deconvolutions contain a lot of data and variables.
'Settings for the deconvolution are stored in cDeconSettings
'Filepaths for the deconvolution files are stored in cDeconFiles
'Deconvoluted contributors are stored as cContributors which contain cLocusSTRmix which contain cGenotypeSTRmix
'***************************************************************************************************

'Private pTargetSheet As String 'Destination for data

'Decon characteristics
Private pSTRmixVersion As String
Private pTimeStamp_DeconXML As String
Private pCaseNum As String
Private pSampleID As String
Private pEvidenceFile As String
Private pVarNOC As Boolean

Private pNOC As Integer
Private pMixProportions As New Scripting.Dictionary 'key = "Contributor #" item = proportion
Private pMixDNA As New Scripting.Dictionary         'key = "Contributor #" item = DNA amount

'Decon file & folder paths
Private pDeconFiles As New cDeconFiles

'Run settings
Private pDeconSettings As New cDeconSettings
Private pSettings_Check As Boolean 'Whether or not it passed the settings check against default

'Run results
Private pIterations As Long
Private pSampleSize As Long
Private pLogLikelihood As Double
Private pGRconvergence As String
Private pGRreject As String
Private pAlleleVariance As Double

'Private pBackStutterVariance As Double
'Private pForwardStutterVariance As Double
'Private pSE33StutterVariance As Double
Private pStutterVarianceResults As New Scripting.Dictionary 'stutter variance result per type of stutter. Key= stutter name, item = stutter variance (double)

Private pDegradation As New Scripting.Dictionary 'Degradation is per contributor

Private pDBTimeStamp As String 'Not the timestamp of the DB search, but the decon. Should match TimeStamp_DeconXML.
Private pDBSearchResults As New Scripting.Dictionary 'from DBSearchResults.xml
Private pDBSearchResultText As String
Private pDatabaseSize As Long 'from DBSearchResults.xml

Private pConditionedList As New Scripting.Dictionary 'conditioned references (Hd knowns)
Private pConditionedOn As String

Private pIgnoredLoci As New Scripting.Dictionary

Private pContributors As New Scripting.Dictionary 'key = contributor #, item = cContributor object
'(cContributors contain cLoci, which contain cGenotypes)

Dim DeconResultsXML As IXMLDOMDocument 'actual xml file loaded as object
Dim DeconConfigXML As IXMLDOMDocument  'actual xml file loaded as object
Dim DBSearchXML As IXMLDOMDocument 'actual xml file loaded as object
Dim rootResults As IXMLDOMElement  'root element of DeconResultsXML
Dim rootConfig As IXMLDOMElement  'root element of DeconConfigXML
Dim rootDBSearch As IXMLDOMElement 'root element of DBSearchXML

Option Explicit


Public Sub InitializeMe(DeconFolder As String, STRmixVersion As String, Optional thisNOC As Integer = 0)
'thisNOC is only specified (<>0) when it's a VarNOC

    'Create cDeconFiles object & grab file/folder paths
    Set Me.DeconFiles = Factory.CreateDeconFiles(DeconFolder, thisNOC, STRmixVersion)
    
    'Import XML files/load into memory
    Call Me.Import_XML
    
    'Grab Decon timestamp & format it
        Dim tempStamp As String
        'STRmix v2.5 uses "runDate", v2.6+ uses "interpretationRunDate"
        If rootResults.selectSingleNode("//interpretationRunDate") Is Nothing Then
            tempStamp = rootResults.selectSingleNode("//runDate").Text
        Else: tempStamp = rootResults.selectSingleNode("//interpretationRunDate").Text
        End If
        
        Me.TimeStamp_DeconXML = Replace((Left(tempStamp, 10) & "-" & Mid(tempStamp, 12, 8)), ":", "-")
        'For a VarNOC, make the timestamp unique:
            If thisNOC <> 0 Then Me.TimeStamp_DeconXML = "(" & thisNOC & ")_" & Me.TimeStamp_DeconXML
    
    'Create cDeconSettings object and harvest settings from DeconConfigXML/rootConfig
    Set Me.DeconSettings = Factory.CreateDeconSettings(rootConfig, Me.TimeStamp_DeconXML, thisNOC)
    
    'Check the DeconSettings against defaults and save as Me.Settings_Check
    Me.Settings_Check = Me.DeconSettings.CheckAgainstDefault()
    
    'Harvest results from DeconResultsXML
    Call Me.HarvestResultsXML(Me.DeconSettings.NOC)
    
    'Harvest stuff from DeconConfigXML, including conditioned contributors
    Call Me.HarvestConfigXML
    
    'Check database search file. Can be turned off on the STRlite Settings tab.
    If Sheets("STRlite Settings").Range("DBSearchOn").Value = True Then Call Me.CheckDBSearch
    
    'Import the ComponentInterp file
    Call Me.ImportContributors
    
    'Create cContributor objects
    Dim i As Integer: i = 1
    Dim newContributor As cContributor
    
        'Create cContributor -> cLocusSTRmix -> cGenotypeSTRmix
        Do While i <= Me.NOC
            Set newContributor = Factory.CreateContributor(Me.TimeStamp_DeconXML, i)
            Me.Contributors.Add newContributor.ContributorName, newContributor
            i = i + 1
        Loop
    
        'Once Contributors are created & initialized, re-name any conditioned cContributors
        'This relies on the fact that conditioned cContributors always come before "unknowns"
        Dim v As Variant
        i = 1
        For Each v In ConditionedList.Keys
            Me.Contributors("Contributor " & i).ContributorName = v
            i = i + 1
        Next v
    
End Sub

Sub Import_XML()
    
    Application.DisplayAlerts = False

    'Import Decon Results xml
    Set DeconResultsXML = CreateObject("MSXML2.DOMDocument")
        With DeconResultsXML
            .async = False
            .validateOnParse = False
            .Load (Me.DeconFiles.FilePath_DeconResultsXML)
        End With
    Set rootResults = DeconResultsXML.DocumentElement
    
    'Import Decon Config xml
    Set DeconConfigXML = CreateObject("MSXML2.DOMDocument")
        With DeconConfigXML
            .async = False
            .validateOnParse = False
            .Load (Me.DeconFiles.FilePath_DeconConfigXML)
        End With
    Set rootConfig = DeconConfigXML.DocumentElement
    
    'Import DBSearchResults xml
    Set DBSearchXML = CreateObject("MSXML2.DOMDocument")
        With DBSearchXML
            .async = False
            .validateOnParse = False
            .Load (Me.DeconFiles.FilePath_SearchResultXML)
        End With
    Set rootDBSearch = DBSearchXML.DocumentElement
      
    Application.DisplayAlerts = True
    
End Sub


Sub HarvestConfigXML()
'Most stuff from the config.xml file goes into cDeconSettings (to be checked against default settings)
'Conditioned contributors (Hd), ignored loci, and the Evidence Filename are also contained in the config file.

    Me.EvidenceFile = Replace(rootConfig.selectSingleNode("//evidence/strmixFile").Text, ".csv", "")
    
    If rootConfig.selectSingleNode("//maxContributors") Is Nothing Then
        Me.VarNOC = False
        Else: Me.VarNOC = True
    End If
    
    'IgnoredLoci and HdKnowns are tags that may have multiple nodes to harvest.
    Dim vNode As Variant 'actually IXMLDOMNode, which is returned in the IXMLDOMNodeLists below
    
    'Get Ignored Loci
        Dim vListIgnored As IXMLDOMNodeList 'returned from .SelectNodes method
        Set vListIgnored = rootConfig.SelectNodes("//ignoredLoci/locus")
        
        For Each vNode In vListIgnored
            Me.IgnoredLoci.Add vNode.Text, Nothing 'fill dictionary (.IgnoredLoci) from IXMLDOMNodeList
        Next vNode
    
    'Get Hd Knowns
        Dim vListConds As IXMLDOMNodeList 'returned from .SelectNodes method
        Set vListConds = rootConfig.SelectNodes("referenceProfiles/reference[@hdKnown=""true""]")
        
        For Each vNode In vListConds
            If Not ConditionedList.Exists(Replace(vNode.Text, "_REF.csv", "")) Then Me.ConditionedList.Add Replace(vNode.Text, "_REF.csv", ""), Nothing 'fill dictionary (.ConditionedList) from IXMLDOMNodeList
        Next vNode
        
        If Me.ConditionedList.Count > 0 Then
            Dim arrCond() As Variant
            arrCond() = Me.ConditionedList.Keys
            Me.ConditionedOn = Join(arrCond, ", ")
        End If

End Sub

Sub HarvestResultsXML(thisNOC As Integer)
'XML tags are case-SENSITIVE so be careful
'thisNOC should <> 0 by now

    Dim newRoot As IXMLDOMElement 'define newRoot as the analysisResult from thisNOC so it works with VarNOC
    Dim i As Integer
    Dim vListProps As Object, vListDNA As Object, vListDegrad As Object
    
    'v2.5 results.xml is totally different. Luckily the results file has the STRmix version #.
    
    Me.STRmixVersion = rootResults.selectSingleNode("//strmixVersion").Text
    
    With Me
    
        If CDbl(Left(Me.STRmixVersion, 3)) < 2.6 Then
            
        'For v2.5
            .CaseNum = rootConfig.selectSingleNode("//caseNumber").Text
            .SampleID = rootConfig.selectSingleNode("//sampleID").Text 'note: tag is "sampleId" in results.xml but "sampleID" in config.xml
            .NOC = IIf(thisNOC = 0, CInt(rootConfig.selectSingleNode("//contributors").Text), thisNOC) 'shouldn't be 0, but just in case
    
            'Rarely, if a decon goes to crap the effective sample size will be "NaN" which I'm guessing means "not a number"?
            If newRoot.selectSingleNode(".//effectiveSampleSize").Text = "NaN" Then
                .SampleSize = 0
            Else: .SampleSize = CLng(newRoot.selectSingleNode(".//effectiveSampleSize").Text)
            End If
            
            .Iterations = CLng(rootResults.selectSingleNode(".//totalIterations").Text)
            .LogLikelihood = CDbl(rootResults.selectSingleNode(".//averageLogLikelihood").Text)
            .GRconvergence = rootResults.selectSingleNode(".//gelmanRubin").Text
            .AlleleVariance = CDbl(rootResults.selectSingleNode(".//results/variance").Text)
    
            'Only one stutter variance in v2.5; saved as both back and forward
            .StutterVarianceResults.Add "Back Stutter", CDbl(rootResults.selectSingleNode(".//stutterVariance").Text)
            .StutterVarianceResults.Add "Forward Stutter", CDbl(rootResults.selectSingleNode(".//stutterVariance").Text)
    
            If rootResults.selectSingleNode(".//rejectedGelmanRubin") Is Nothing Then
                Me.GRreject = ""
                Else: Me.GRreject = rootResults.selectSingleNode(".//rejectedGelmanRubin").Text
            End If
    
            'Get Lists of mixture proportions & DNA amounts for contributors
            Set vListProps = rootResults.SelectNodes(".//mixtureProportion")
            Set vListDNA = rootResults.SelectNodes(".//dnaAmount")
            
            For i = 1 To Me.NOC
                If Not MixProportions.Exists("Contributor " & i) Then Me.MixProportions.Add "Contributor " & i, CDbl(vListProps.Item(i - 1).Text) / 100
                If Not MixDNA.Exists("Contributor " & i) Then Me.MixDNA.Add "Contributor " & i, CLng(vListDNA.Item(i - 1).Text)
                If Not Degradation.Exists("Contributor " & i) Then Me.Degradation.Add "Contributor " & i, CDbl(0) 'set the non-existent d-value to 0
            Next
    
        Else:
        
        'For v2.6+
    
            'Case num & sample ID are independent of NOC so we can use rootResults
            .CaseNum = rootResults.selectSingleNode("//caseNumber").Text
            .SampleID = rootResults.selectSingleNode("//sampleId").Text 'note: tag is "sampleId" in results.xml but "sampleID" in config.xml
            .NOC = IIf(thisNOC = 0, CInt(rootResults.selectSingleNode("//numberOfContributors/contributors").Text), thisNOC) 'shouldn't be 0, but just in case
            
            Set newRoot = rootResults.selectSingleNode("//analysisResult[@numberOfContributors='" & CStr(Me.NOC) & "']")
            'The rest of the tags could be NOC-dependent in varNOC. From here on, use Me.NOC instead of thisNOC
            
            'Rarely, if a decon goes to crap the effective sample size will be "NaN"
            .SampleSize = IIf(newRoot.selectSingleNode(".//effectiveSampleSize").Text = "NaN", 0, CLng(newRoot.selectSingleNode(".//effectiveSampleSize").Text))
            
            .Iterations = CLng(newRoot.selectSingleNode(".//totalIterations").Text)
            .LogLikelihood = CDbl(newRoot.selectSingleNode(".//averageLogLikelihood").Text)
            .GRconvergence = newRoot.selectSingleNode(".//gelmanRubin").Text
            .AlleleVariance = CDbl(newRoot.selectSingleNode(".//results/variance").Text)   'watch out for stutter variances- similar tags
            
            'Get stutter variances
            Dim vListVariances As IXMLDOMNodeList
            Dim strStutter As String
            
            Set vListVariances = newRoot.SelectNodes(".//stutterVariances/variance")
            Dim v As IXMLDOMElement
            For Each v In vListVariances
                If Not StutterVarianceResults.Exists(v.Attributes(0).Text) Then
                    StutterVarianceResults.Add v.Attributes(0).Text, CDbl(v.Text)
                End If
            Next v
            
            
'            .BackStutterVariance = CDbl(newRoot.SelectSingleNode(".//variance[@stutter='Back Stutter']").Text)
'            .ForwardStutterVariance = CDbl(newRoot.SelectSingleNode(".//variance[@stutter='Forward Stutter']").Text)
'
'            If newRoot.SelectSingleNode(".//variance[@stutter='SE33 -2 Stutter']") Is Nothing Then
'                .SE33StutterVariance = ""
'                Else: .SE33StutterVariance = CDbl(newRoot.SelectSingleNode(".//variance[@stutter='SE33 -2 Stutter']").Text)
'            End If
            
            
            'Gelman-rubin
            If newRoot.selectSingleNode(".//rejectedGelmanRubin") Is Nothing Then
                Me.GRreject = ""
                Else: Me.GRreject = newRoot.selectSingleNode(".//rejectedGelmanRubin").Text
            End If
        
            'Get Lists of mixture proportions & DNA amounts for contributors
            Set vListProps = newRoot.SelectNodes(".//mixtureProportion")
            Set vListDNA = newRoot.SelectNodes(".//dnaAmount")
            Set vListDegrad = newRoot.SelectNodes(".//contributorDegradationCurveDValue")
            
            For i = 1 To Me.NOC
                If Not MixProportions.Exists("Contributor " & i) Then Me.MixProportions.Add "Contributor " & i, CDbl(vListProps.Item(i - 1).Text) / 100
                If Not MixDNA.Exists("Contributor " & i) Then Me.MixDNA.Add "Contributor " & i, CLng(vListDNA.Item(i - 1).Text)
                If Not Degradation.Exists("Contributor " & i) Then Me.Degradation.Add "Contributor " & i, CDbl(vListDegrad.Item(i - 1).Text)
            Next
    
        End If
    
    End With
    


End Sub


Sub ImportContributors()
'The Import tab is used temporarily to store & sort locus/allele data from ComponentInterp,
'so it is cleared & reused with each Decon.

    Dim wkbTemp As Workbook
    Dim CurrentFile As String: CurrentFile = ThisWorkbook.Name
    
    With Workbooks(CurrentFile).Sheets("Import")
        .Visible = True
        .Cells.ClearContents
    End With

    With Workbooks.Open(Me.DeconFiles.FilePath_Contributor)
        .Sheets(1).UsedRange.Copy destination:=Workbooks(CurrentFile).Sheets("Import").Range("A1")
        .Close False
    End With

'    Set wkbTemp = Workbooks.Open(Me.DeconFiles.FilePath_Contributor)
'    wkbTemp.Activate
'    wkbTemp.Sheets(1).Range("A:A").Copy
'
'    Workbooks(CurrentFile).Sheets("Import").Activate
'    Workbooks(CurrentFile).Sheets("Import").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'    :=False, Transpose:=False
    
    Application.CutCopyMode = False
'    wkbTemp.Close

End Sub

Sub CheckDBSearch()

If Me.DeconFiles.FilePath_SearchResultXML = "n/a" Then 'If the DBSearch file was never found...
    
    Me.DBSearchResultText = "Check Report"
    
Else:

    If rootDBSearch.selectSingleNode("//interpretationRunDate") Is Nothing Then 'v2.5 calls it "runDate"
        Me.DBTimeStamp = rootDBSearch.selectSingleNode("//runDate").Text
        Else: Me.DBTimeStamp = rootDBSearch.selectSingleNode("//interpretationRunDate").Text
    End If
    
    Me.DatabaseSize = CLng(rootDBSearch.selectSingleNode("//numberInDatabase").Text)

    Dim v As Variant, vNode As Variant 'actually IXMLDOMNode, which is returned in the IXMLDOMNodeLists below
    Dim vListDB As Object 'actually IXMLDOMNodeList, returned from .SelectNodes method
        Set vListDB = rootDBSearch.SelectNodes("//stdResult")
        
        For Each vNode In vListDB
            'fill dictionary (.DBSearchResults) key = caseNumber_sample (attributes 1 & 2 of that node), item = LR (child node)
            If Not DBSearchResults.Exists(vNode.Attributes(0).Text & "_" & vNode.Attributes(1).Text) Then Me.DBSearchResults.Add vNode.Attributes(0).Text & "_" & vNode.Attributes(1).Text, vNode.FirstChild.Text
        Next vNode
    
    If DBSearchResults.Count > 0 Then
        Dim arrResults() As Variant
        Dim templist As String: templist = ""
        arrResults = Me.DBSearchResults.Keys
        Me.DBSearchResultText = Join(arrResults, ", ")
    
        For Each v In Me.DBSearchResults.Keys
            templist = templist & v & vbTab & "LR: " & Me.DBSearchResults(v) & vbNewLine
        Next v
        
        MsgBox Me.CaseNum & "_" & Me.SampleID & vbNewLine & vbNewLine & _
            "Database Search Results: " & vbNewLine & vbNewLine & _
            templist, vbExclamation + vbOKOnly, "STRmix Database Hit"

    End If
    
End If

End Sub


Public Sub DumpData(TargetSheet As Worksheet, SheetType As String)

    Application.ScreenUpdating = False
    TargetSheet.Activate
    
    Dim DestTimeStamp As Range
    Dim DestResults As Range
    Dim DestConfig As Range
    Dim DestFiles As Range
    Dim DestMix As Range
    Dim DestGraphs As Range
    Dim DestComponents As Range
    Dim DestLocusCat As Range
    Dim DestSpecCat As Range
    
    Dim i As Integer, rng As Range, v As Variant
    
    
    'Check for VarNOC
    If Me.VarNOC = True And SheetType <> "Decon" Then
        MsgBox Me.CaseNum & "_" & Me.SampleID & vbNewLine & vbNewLine & _
        "This VarNOC sample will be saved as a regular Deconvolution Worksheet.", vbExclamation + vbOKOnly, "VarNOC"
        SheetType = "Decon"
    End If
    
    'Check NOC to make sure
    If SheetType = "1P" And Me.NOC <> 1 Then
        MsgBox Me.CaseNum & "_" & Me.SampleID & vbNewLine & vbNewLine & _
        "This sample cannot fit on the 1P Worksheet." & vbNewLine & _
        "It will be saved as a regular Deconvolution Worksheet.", vbExclamation + vbOKOnly, "Wrong NOC!"
        SheetType = "Decon"
        'Create new Decon tab
            Set TargetSheet = Factory.CreateDeconSheet(Me)
    End If

    If SheetType = "2P" And Me.NOC <> 2 Then
        MsgBox Me.CaseNum & "_" & Me.SampleID & vbNewLine & vbNewLine & _
        "This sample cannot fit on the 2P Worksheet." & vbNewLine & _
        "It will be saved as a regular Deconvolution Worksheet.", vbExclamation + vbOKOnly, "Wrong NOC!"
        SheetType = "Decon"
        'Create new Decon tab
            Set TargetSheet = Factory.CreateDeconSheet(Me)
    End If

    

    Select Case SheetType
    
        Case "Decon"
        
            Set DestTimeStamp = TargetSheet.Range("DeconTimestamp")
            Set DestResults = TargetSheet.Range("Dest_DeconResults")
            Set DestConfig = TargetSheet.Range("Dest_DeconConfig")
            Set DestFiles = TargetSheet.Range("Dest_DeconFiles")
            Set DestMix = TargetSheet.Range("Dest_MixProp")
            Set DestGraphs = TargetSheet.Range("O2")
            Set DestComponents = TargetSheet.Range("Dest_Components")
            Set DestLocusCat = TargetSheet.Range("Dest_LocusCat")
            Set DestSpecCat = TargetSheet.Range("Dest_SpecimenCategory")
            
        Case "1P"
        
            Set DestTimeStamp = TargetSheet.Range("DeconTimestamp1P")
            Set DestResults = TargetSheet.Range("Dest_DeconResults1P")
            Set DestConfig = TargetSheet.Range("Dest_DeconConfig1P")
            Set DestFiles = TargetSheet.Range("Dest_DeconFiles1P")
            Set DestMix = TargetSheet.Range("Dest_MixProp1P")
            Set DestGraphs = TargetSheet.Range("N2")
            Set DestComponents = TargetSheet.Range("Dest_1PComponents")
            Set DestLocusCat = TargetSheet.Range("Dest_1PLocusCat")
            Set DestSpecCat = TargetSheet.Range("Dest_1PSpecCat")
        
        Case "2P"

            Set DestTimeStamp = TargetSheet.Range("DeconTimestamp2P")
            Set DestResults = TargetSheet.Range("Dest_DeconResults2P")
            Set DestConfig = TargetSheet.Range("Dest_DeconConfig2P")
            Set DestFiles = TargetSheet.Range("Dest_DeconFiles2P")
            Set DestMix = TargetSheet.Range("Dest_MixProp2P")
            Set DestGraphs = TargetSheet.Range("N2")
            Set DestComponents = TargetSheet.Range("Dest_2PComponents")
            Set DestLocusCat = TargetSheet.Range("Dest_2PLocusCat")
            Set DestSpecCat = TargetSheet.Range("Dest_2PSpecCat")
    
    End Select

    DestTimeStamp.Value = Me.TimeStamp_DeconXML
    
    'Decon Results
    With DestResults
        .Offset(0, 0) = Me.CaseNum
        .Offset(1, 0) = Me.SampleID
        .Offset(2, 0) = Me.EvidenceFile
        .Offset(3, 0) = Me.NOC
        
        .Offset(4, 0) = Me.Iterations
        .Offset(5, 0) = Me.SampleSize
        .Offset(6, 0) = Me.LogLikelihood
        .Offset(7, 0) = Me.GRconvergence
        .Offset(8, 0) = Me.GRreject

        .Offset(9, 0) = Me.AlleleVariance
        
        i = 1
        For Each v In Me.StutterVarianceResults.Keys
            .Offset(9 + i, -1) = v & " Variance" 'stutter name
            .Offset(9 + i, 0) = Me.StutterVarianceResults(v) 'stutter variance
            i = i + 1
        Next v
        
    End With
    
    'Decon config settings
    Dim config As cDeconSettings: Set config = Me.DeconSettings 'Shortcut for Me.DeconSettings object
    Dim IgnoredList As String, arrIgnored() As Variant
        arrIgnored() = Me.IgnoredLoci.Keys
        IgnoredList = Join(arrIgnored, ", ")
        
    With DestConfig
        .Offset(0, 0) = config.CaseNum
        .Offset(1, 0) = config.SampleID
        .Offset(2, 0) = config.NOC
        .Offset(3, 0) = config.BurnAccepts
        .Offset(4, 0) = config.PostBurnAccepts
        .Offset(5, 0) = config.Chains
        .Offset(6, 0) = config.RWSD
        .Offset(7, 0) = config.PostBurnShortlist
        .Offset(8, 0) = config.MxPriors
        .Offset(9, 0) = config.AutoContinueGR
        .Offset(10, 0) = config.GRthreshold
        .Offset(11, 0) = config.ExtraAccepts
        .Offset(12, 0) = config.KitName
        .Offset(13, 0) = config.AutosomalLoci
        .Offset(14, 0) = config.Saturation
        .Offset(15, 0) = config.DegradStart
        .Offset(16, 0) = config.DegradMax
        .Offset(17, 0) = config.DropInCap
        .Offset(18, 0) = config.DropInFreq
        .Offset(19, 0) = config.DropInGamma
        .Offset(20, 0) = config.AlleleVariance
        .Offset(21, 0) = config.MinVarianceFactor
        .Offset(22, 0) = config.LocusAmpVariance
        
        i = 0
        For Each v In config.StutterSettings
            .Offset(23 + (2 * i), -1) = v & " Max"
            .Offset(23 + (2 * i), 0) = config.StutterSettings(v).StutterMax
            .Offset(24 + (2 * i), -1) = v & " Variance"
            .Offset(24 + (2 * i), 0) = config.StutterSettings(v).StutterVariance
            i = i + 1
        Next v
        
        .Offset(31, 0) = IIf(Me.Settings_Check, "Default", "Check")
        .Offset(32, 0) = Me.ConditionedList.Count 'integer
        .Offset(33, 0) = Me.ConditionedOn 'string
        .Offset(34, 0) = IgnoredList

    'DB search Results
        .Offset(37, 0) = Me.DBTimeStamp
        .Offset(38, 0) = Me.DatabaseSize
        .Offset(39, 0) = Me.DBSearchResultText
    
    End With
              
    'Files & Folders
    
    With DestFiles
        .Offset(0, 0) = Me.DeconFiles.DeconFolderBasename
        .Offset(1, 0) = Me.DeconFiles.FolderName_Decon
        .Offset(2, 0) = Me.DeconFiles.FolderPath_Decon
        .Offset(3, 0) = Me.DeconFiles.FileName_Decon
        .Offset(4, 0) = Me.DeconFiles.FilePath_Decon
        .Offset(5, 0) = Me.DeconFiles.FilePath_Contributor
        .Offset(6, 0) = Me.DeconFiles.FilePath_DeconConfigXML
        .Offset(7, 0) = Me.DeconFiles.FilePath_DeconResultsXML
        .Offset(8, 0) = Me.DeconFiles.FolderPath_Parent
        .Offset(9, 0) = Me.DeconFiles.FolderName_Parent
        .Offset(10, 0) = Me.DeconFiles.FolderPath_Reports
        .Offset(11, 0) = Me.DeconFiles.FilePath_AlleleChart
        'Offset(12, 0) = spot for Filepath_DegCurveChart maybe
        
        'Stutter Variance Charts
        i = 0
        For Each v In Me.DeconSettings.StutterSettings
            .Offset(13 + i, -1) = v & " Chart FilePath"
            .Offset(13 + i, 0) = Me.DeconFiles.FilePath_StutterCharts(Me.DeconSettings.StutterSettings(v).StutterChart)
            i = i + 1
        Next v
        
    End With
    
    'Contributor proportions & DNA amounts
    For i = 0 To Me.NOC - 1
        DestMix.Offset(i, 1) = Me.MixProportions.Items(i) 'Mix proportion
        DestMix.Offset(i, 2) = vbCr & Me.MixDNA.Items(i) & " RFU" 'DNA amount, with a carriage return for the pie chart legend
        DestMix.Offset(i, 3) = Me.Degradation.Items(i) 'degradation
        DestMix.Offset(i, 4) = Me.MixDNA.Keys(i) 'Contributor Name Only
        
        'Degradation flag!!
        If Me.Degradation.Items(i) > DegradationFlag Then
            MsgBox "High contributor degradation detected:" & vbNewLine & vbNewLine & _
                    Me.CaseNum & "_" & Me.SampleID & vbNewLine & "Contributor " & i + 1 & ":  " & Me.Degradation.Items(i), vbExclamation + vbOKOnly, _
                    "High degradation detected"
        End If
    Next i
    
    'Conditioned names
    Dim cont As Variant
    i = 0
    For Each cont In Me.ConditionedList.Keys
        DestMix.Offset(i, 4) = cont 'replaces the Contributors in the Contributor Name column
        i = i + 1
    Next cont
    
    'Contributor Summary (Components>threshold)
    Dim Locus As Range
    Call Admin.LociCollection
    TargetSheet.Activate
    For Each Locus In DestComponents
        If LociSTRmix.Exists(Locus.Value) Then 'if the locus is included in the STRmix set
            'If it's an Ignored Locus, then flag it across all contributors
            If Me.IgnoredLoci.Exists(Locus.Value) Then
                For i = 1 To Me.NOC
                    If SheetType = "Decon" Then
                        Locus.Offset(0, (2 * i) - 1).Value = "Omitted"
                    Else:
                        Locus.Offset(0, i).Value = "Omitted"
                    End If
                    
                Next i
            Else:
            'If locus is in cContributor->cLocus->LociPulled list, retrieve FinalPull
                i = 1 'contributor count
                For Each cont In Me.Contributors.Items
                    If cont.LociPulled.Exists(Locus.Value) Then
                        If SheetType = "Decon" Then
                            Locus.Offset(0, (2 * i) - 1).Value = cont.LociPulled(Locus.Value).FinalPull
                        Else:
                            Locus.Offset(0, i).Value = cont.LociPulled(Locus.Value).FinalPull
                        End If
                    End If
                    i = i + 1
                Next cont
            End If
        End If
    Next Locus
    
    'Locus Specimen categories
    For Each Locus In DestLocusCat
        If LociSTRmix.Exists(Locus.Value) Then 'if the locus is included in the STRmix set
            'If locus is in cContributor->cLocus->LociPulled list, retrieve LocusCategory
                i = 1 'contributor count
                For Each cont In Me.Contributors.Items
                    If cont.LociPulled.Exists(Locus.Value) Then Locus.Offset(0, i).Value = IIf(cont.LociAll(Locus.Value).LocusCategory <> "U", "M", "")
                    i = i + 1
                Next cont
        End If
    Next Locus
    
    'Overall specimen categories & MME/MRE (per contributor)
    i = 1
    For Each cont In Me.Contributors.Items
        DestSpecCat.Offset(0, i).Value = cont.SpecimenCategory
        DestSpecCat.Offset(1, i).Value = cont.MME
        DestSpecCat.Offset(2, i).Value = cont.SDIS_MRE
        DestSpecCat.Offset(3, i).Value = cont.NDIS_MRE
        i = i + 1
    Next cont
    
    'Variance Graphs
    
    On Error Resume Next
    Application.DisplayAlerts = False
    
    Dim shp As Shape
    For Each shp In TargetSheet.Shapes
        If shp.Type = msoPicture Then shp.Delete
    Next shp
    
    Call InsertGraph(TargetSheet, DestGraphs, Me.DeconFiles.FilePath_AlleleChart, "AlleleVariance")
    
    If Left(Me.STRmixVersion, 3) = "2.5" Then
        Call InsertGraph(TargetSheet, DestGraphs.Offset(15, 0), Me.DeconFiles.FilePath_StutterCharts("Back Stutter"), "StutterVariance")
    Else:
        'For v2.6+
        i = 1
    
        For Each v In Me.DeconSettings.StutterSettings
        Debug.Print Me.DeconSettings.StutterSettings(v).StutterChart
        Debug.Print Me.DeconFiles.FilePath_StutterCharts(Me.DeconSettings.StutterSettings(v).StutterChart)
            Select Case i
                Case Is < 3
                    Call InsertGraph(TargetSheet, DestGraphs.Offset(15, (i - 1) * 2), Me.DeconFiles.FilePath_StutterCharts(Me.DeconSettings.StutterSettings(v).StutterChart), v & " Variance")
                Case 3
                    Call InsertGraph(TargetSheet, DestGraphs.Offset(0, 2), Me.DeconFiles.FilePath_StutterCharts(Me.DeconSettings.StutterSettings(v).StutterChart), v & " Variance")
                Case 4
                    Call InsertGraph(TargetSheet, DestGraphs.Offset(30, 0), Me.DeconFiles.FilePath_StutterCharts(Me.DeconSettings.StutterSettings(v).StutterChart), v & " Variance")
            End Select
            i = i + 1
        Next v
        
    End If
    
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    
    Select Case SheetType
        Case "Decon"
        
            'Mixture pie chart- only on Decon sheet
            ActiveWorkbook.ChartDataPointTrack = False
            
            Dim ChartData As Range
                
            If Me.NOC > 0 Then
                Set ChartData = TargetSheet.Range("Dest_MixProp").Resize(Me.NOC, 2)
            Else
                Set ChartData = TargetSheet.Range("V7:W7")
            End If
        
            TargetSheet.ChartObjects("Chart_MixProp").Chart.SetSourceData Source:=ChartData, PlotBy:=xlColumns
        
            'Adjust "Send to CODIS" buttons
            'Turn off all buttons
            For i = 1 To 4
                TargetSheet.OLEObjects("ToCODIS" & i).Visible = False
                TargetSheet.OLEObjects("CondtoCODIS" & i).Visible = False
            Next i
        
            'Turn on all regular buttons up to Contributor #
            If Me.NOC <> 0 Then
                For i = 1 To Me.NOC
                    TargetSheet.OLEObjects("ToCODIS" & i).Visible = True
                Next i
            End If
        
            If Me.ConditionedList.Count <> 0 Then
                For i = 1 To Me.ConditionedList.Count
                    TargetSheet.OLEObjects("ToCODIS" & i).Visible = False
                    TargetSheet.OLEObjects("CondtoCODIS" & i).Visible = True
                Next i
            End If
            
        Case "1P"
            TargetSheet.OLEObjects("ToCODIS1").Visible = IIf(Application.WorksheetFunction.CountA(TargetSheet.Range("Dest_1PComponents").Offset(0, 1)) = 0, False, True)
        
        Case "2P"
            TargetSheet.OLEObjects("ToCODIS1").Visible = IIf(Application.WorksheetFunction.CountA(TargetSheet.Range("Dest_2PLociPreSTRmix").Offset(0, 4)) = 0, False, True)
            TargetSheet.OLEObjects("ToCODIS2").Visible = IIf(Application.WorksheetFunction.CountA(TargetSheet.Range("Dest_2PLociPreSTRmix").Offset(0, 5)) = 0, False, True)
            TargetSheet.OLEObjects("CondtoCODIS1").Visible = IIf(Me.ConditionedList.Count > 0, True, False)
            
    End Select
    
    Application.ScreenUpdating = True
    
End Sub

Sub InsertGraph(TargetSheet As Worksheet, TargetCell As Range, FilePath As String, GraphName As String)

        With TargetSheet.Pictures.insert(FilePath)
        With .ShapeRange
            .LockAspectRatio = msoFalse
            .Width = 288
            .Height = 216
            .Name = GraphName
        End With
        .Left = TargetCell.Left
        .Top = TargetCell.Top
        .Placement = 1
        .PrintObject = False
    End With

End Sub


'****************************************************
'                cDecon Properties
'****************************************************

Public Property Get STRmixVersion() As String
    STRmixVersion = pSTRmixVersion
End Property
Public Property Let STRmixVersion(v As String)
    pSTRmixVersion = v
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

Public Property Get EvidenceFile() As String
    EvidenceFile = pEvidenceFile
End Property
Public Property Let EvidenceFile(v As String)
    pEvidenceFile = v
End Property

Public Property Get VarNOC() As Boolean
    VarNOC = pVarNOC
End Property
Public Property Let VarNOC(v As Boolean)
    pVarNOC = v
End Property

Public Property Get NOC() As Integer
    NOC = pNOC
End Property
Public Property Let NOC(v As Integer)
    pNOC = v
End Property

Public Property Get ConditionedOn() As String
    ConditionedOn = pConditionedOn
End Property
Public Property Let ConditionedOn(v As String)
    pConditionedOn = v
End Property

Public Property Get ConditionedList() As Dictionary
    Set ConditionedList = pConditionedList
End Property
Public Property Let ConditionedList(v As Dictionary)
    Set pConditionedList = v
End Property

Public Property Get IgnoredLoci() As Dictionary
    Set IgnoredLoci = pIgnoredLoci
End Property
Public Property Let IgnoredLoci(v As Dictionary)
    Set pIgnoredLoci = v
End Property

Public Property Get MixProportions() As Dictionary
    Set MixProportions = pMixProportions
End Property
Public Property Let MixProportions(v As Dictionary)
    Set pMixProportions = v
End Property

Public Property Get MixDNA() As Dictionary
    Set MixDNA = pMixDNA
End Property
Public Property Let MixDNA(v As Dictionary)
    Set pMixDNA = v
End Property


Public Property Get TimeStamp_DeconXML() As String
    TimeStamp_DeconXML = pTimeStamp_DeconXML
End Property
Public Property Let TimeStamp_DeconXML(v As String)
    pTimeStamp_DeconXML = v
End Property



Public Property Get DBTimeStamp() As String
    DBTimeStamp = pDBTimeStamp
End Property
Public Property Let DBTimeStamp(v As String)
    pDBTimeStamp = v
End Property

Public Property Get DBSearchResults() As Dictionary
    Set DBSearchResults = pDBSearchResults
End Property
Public Property Let DBSearchResults(v As Dictionary)
    Set pDBSearchResults = v
End Property

Public Property Get DBSearchResultText() As String
    DBSearchResultText = pDBSearchResultText
End Property
Public Property Let DBSearchResultText(v As String)
    pDBSearchResultText = v
End Property

Public Property Get DatabaseSize() As Long
    DatabaseSize = pDatabaseSize
End Property
Public Property Let DatabaseSize(v As Long)
    pDatabaseSize = v
End Property


'File/folder paths
Public Property Get DeconFiles() As cDeconFiles
    Set DeconFiles = pDeconFiles
End Property
Public Property Set DeconFiles(v As cDeconFiles)
    Set pDeconFiles = v
End Property



'Diagnostics

Public Property Get Iterations() As Long
    Iterations = pIterations
End Property
Public Property Let Iterations(v As Long)
    pIterations = v
End Property

Public Property Get SampleSize() As Long
    SampleSize = pSampleSize
End Property
Public Property Let SampleSize(v As Long)
    pSampleSize = v
End Property

Public Property Get LogLikelihood() As Double
    LogLikelihood = pLogLikelihood
End Property
Public Property Let LogLikelihood(v As Double)
    pLogLikelihood = v
End Property

Public Property Get GRconvergence() As String
    GRconvergence = pGRconvergence
End Property
Public Property Let GRconvergence(v As String)
    pGRconvergence = v
End Property

Public Property Get GRreject() As String
    GRreject = pGRreject
End Property
Public Property Let GRreject(v As String)
    pGRreject = v
End Property

Public Property Get AlleleVariance() As Double
    AlleleVariance = pAlleleVariance
End Property
Public Property Let AlleleVariance(v As Double)
    pAlleleVariance = v
End Property


Public Property Get StutterVarianceResults() As Dictionary
    Set StutterVarianceResults = pStutterVarianceResults
End Property
Public Property Let StutterVarianceResults(v As Dictionary)
    Set pStutterVarianceResults = v
End Property

Public Property Get Degradation() As Dictionary
    Set Degradation = pDegradation
End Property
Public Property Let Degradation(v As Dictionary)
    Set pDegradation = v
End Property



Public Property Get Contributors() As Dictionary
    Set Contributors = pContributors
End Property
Public Property Let Contributors(v As Dictionary)
    Set pContributors = v
End Property


'Setting Checks

Public Property Get Settings_Check() As String
    Settings_Check = pSettings_Check
End Property
Public Property Let Settings_Check(v As String)
    pSettings_Check = v
End Property

Public Property Get DeconSettings() As cDeconSettings
    Set DeconSettings = pDeconSettings
End Property
Public Property Set DeconSettings(v As cDeconSettings)
    Set pDeconSettings = v
End Property

