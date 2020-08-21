VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDeconSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pTimeStampDecon As String
Private pSTRmixVersion As String

Private pCaseNum As String
Private pSampleID As String
Private pNOC As Integer

Private pBurnAccepts As Long
Private pPostBurnAccepts As Long
Private pChains As Integer
Private pRWSD As Double
Private pPostBurnShortlist As Integer
Private pMxPriors As Boolean
Private pAutoContinueGR As Boolean
Private pGRthreshold As Double
Private pExtraAccepts As Long

Private pHRpercentAccepts As Double 'varNOC
Private pPopForRange As String 'varNOC

Private pKitName As String
Private pAutosomalLoci As Integer
Private pSaturation As Long
Private pDegradStart As Double
Private pDegradMax As Double
Private pDropInCap As Integer
Private pDropInFreq As Double
Private pDropInGamma As String
Private pAlleleVariance As String
Private pMinVarianceFactor As Double
Private pLocusAmpVariance As Double
Private pVarMinParameter As Double

Private pStutterSettings As New Scripting.Dictionary 'key = stutter name, item = cStutterModel object

'Private pBackStutterMax As Double
'Private pForwardStutterMax As Double
'Private pSE33StutterMax As Double
'Private pBackStutterVariance As String
'Private pForwardStutterVariance As String
'Private pSE33StutterVariance As String

Private pAutoSearch As Boolean

Public Sub InitializeMe(rootConfig As Object, TimeStampID As String, thisNOC As Integer)
'rootConfig = IXMLDOMElement (root element of DeconConfigXML)
'the TimeStampID is to label this particular cDeconSettings object. It comes from the sister decon file, Results.xml
'thisNOC will be 0 unless it's a VarNOC

    'Harvest all the config settings from rootConfig. Most are single unique nodes (tags).
    With Me
        
        .TimeStampDecon = TimeStampID
        .CaseNum = CStr(rootConfig.selectSingleNode("//caseNumber").Text)
        .SampleID = CStr(rootConfig.selectSingleNode("//sampleID").Text) 'note: "sampleId" in results.xml but "sampleID" in config.xml
        
        'Handle VarNOC (thisNOC<>0)
        If thisNOC <> 0 Then
            .HRpercentAccepts = CDbl(rootConfig.selectSingleNode("//hrPercentAccepts").Text) 'VarNOC setting
            .PopForRange = rootConfig.selectSingleNode("//popForRange").Text 'VarNOC setting
            .NOC = thisNOC
        Else:
            .HRpercentAccepts = 0
            .PopForRange = ""
            .NOC = CInt(rootConfig.selectSingleNode("//contributors").Text)
        End If

        
        .BurnAccepts = CLng(rootConfig.selectSingleNode("//burninAccepts").Text)
        .PostBurnAccepts = CLng(rootConfig.selectSingleNode("//postBurninAccepts").Text)
        .Chains = CInt(rootConfig.selectSingleNode("//chains").Text)
        .RWSD = CDbl(rootConfig.selectSingleNode("//rwsd").Text)
        .PostBurnShortlist = CInt(rootConfig.selectSingleNode("//postBurninShortlist").Text)
        .MxPriors = CBool(rootConfig.selectSingleNode("//useMxPriors").Text)
        .AutoContinueGR = CBool(rootConfig.selectSingleNode("//autoContinue").Text)
        
        If .AutoContinueGR Then
            .GRthreshold = CDbl(rootConfig.selectSingleNode("//gelmanRubinThreshold").Text)
            If rootConfig.selectSingleNode("//extraAccepts") Is Nothing Then 'v2.6+ calls it "extraAccepts"
                .ExtraAccepts = CLng(rootConfig.selectSingleNode("//extraIterations").Text) 'v2.5 calls it "extraIterations"
            Else: .ExtraAccepts = CLng(rootConfig.selectSingleNode("//extraAccepts").Text)
            End If
        Else:
            .GRthreshold = 0
            .ExtraAccepts = 0
        End If
        

        .KitName = rootConfig.selectSingleNode("//kits/kit/@name").Text
        
        .AutosomalLoci = CInt(rootConfig.selectSingleNode("//numNonGenderLoci").Text)
        .Saturation = CLng(rootConfig.selectSingleNode("//saturation").Text)
        .DegradStart = CDbl(rootConfig.selectSingleNode("//degradationStartsAt").Text)
        .DegradMax = CDbl(rootConfig.selectSingleNode("//degradationMax").Text)
        .DropInCap = CInt(rootConfig.selectSingleNode("//dropinCap").Text)
        .DropInFreq = CDbl(rootConfig.selectSingleNode("//dropinFrequency").Text)
        .DropInGamma = rootConfig.selectSingleNode("//dropinGammaParameters").Text
        .AlleleVariance = rootConfig.selectSingleNode("//allelicVariance").Text
        .MinVarianceFactor = CDbl(rootConfig.selectSingleNode("//minVarFactor").Text)
        .LocusAmpVariance = CDbl(rootConfig.selectSingleNode("//locusAmpVariance").Text)
        .AutoSearch = CBool(rootConfig.selectSingleNode("//autoSearch").Text)
        
        If rootConfig.selectSingleNode("//varianceMinimizationParameter") Is Nothing Then 'v2.5 doesn't have this
            .VarMinParameter = 0
        Else: .VarMinParameter = CDbl(rootConfig.selectSingleNode("//varianceMinimizationParameter").Text)
        End If
        
        
    'Set stutter max/variances. Need to handle v2.5 too.
    'I wish that the config file had the STRmix version # in it
    
        Dim newStutter As cStutterModel
        
        If rootConfig.selectSingleNode("//stutters") Is Nothing Then
        
            Me.STRmixVersion = "2.5"
            'v2.5 doesn't have the node "//stutters". It has only one stutter variance and two stutter maxes (forward and back)
                Dim backStutter As New cStutterModel
                Dim forwardStutter As New cStutterModel
                With backStutter
                    .StutterName = "Back Stutter"
                    .StutterMax = CDbl(rootConfig.selectSingleNode("//stutterMax").Text)
                    .StutterVariance = CStr(rootConfig.selectSingleNode("//stutterVariance").Text)
                End With
                With forwardStutter
                    .StutterName = "Forward Stutter"
                    .StutterMax = CDbl(rootConfig.selectSingleNode("//forwardStutterMax").Text)
                    .StutterVariance = CStr(rootConfig.selectSingleNode("//stutterVariance").Text)
                End With
                
                If Not StutterSettings.Exists(backStutter.StutterName) Then StutterSettings.Add backStutter.StutterName, backStutter
                If Not StutterSettings.Exists(forwardStutter.StutterName) Then StutterSettings.Add forwardStutter.StutterName, forwardStutter
            
        Else:
            
            Me.STRmixVersion = "2.6+"
            'v2.6+ has separate stutter settings for each type of stutter
            Dim v As IXMLDOMElement, i As Integer
            Dim StutterList As IXMLDOMNodeList: Set StutterList = rootConfig.SelectNodes("//stutters/stutter")
            
            i = 1
            For Each v In StutterList
                If Not StutterSettings.Exists(v.Attributes(0).Text) Then
                    Set newStutter = New cStutterModel
                    newStutter.StutterName = v.Attributes(0).Text
                    newStutter.StutterEnabled = IIf(v.selectSingleNode("enabled").Text = "Y", True, False)
                    newStutter.StutterMax = CDbl(v.selectSingleNode("stutterMax").Text)
                    newStutter.StutterVariance = CStr(v.selectSingleNode("stutterVariance").Text)
                    newStutter.StutterChart = "Var_Stut" & i
                    StutterSettings.Add newStutter.StutterName, newStutter
                    i = i + 1
                End If
            Next v
            
        End If
        
'            .BackStutterMax = CDbl(rootConfig.selectSingleNode("//stutter[@name= ""Back Stutter""]/stutterMax").Text)
'            .BackStutterVariance = CStr(rootConfig.selectSingleNode("//stutter[@name= ""Back Stutter""]/stutterVariance").Text)
'            .ForwardStutterMax = CDbl(rootConfig.selectSingleNode("//stutter[@name= ""Forward Stutter""]/stutterMax").Text)
'            .ForwardStutterVariance = CStr(rootConfig.selectSingleNode("//stutter[@name= ""Forward Stutter""]/stutterVariance").Text)
            
'            'If the lab isn't modeling SE33 stutter (Fusion 5):
'            If rootConfig.selectSingleNode("//stutter[@name= ""SE33 -2 Stutter""]") Is Nothing Then
'                .SE33StutterMax = 0
'                .SE33StutterVariance = "n/a"
'            Else:
'                .SE33StutterMax = CDbl(rootConfig.selectSingleNode("//stutter[@name= ""SE33 -2 Stutter""]/stutterMax").Text)
'                .SE33StutterVariance = CStr(rootConfig.selectSingleNode("//stutter[@name= ""SE33 -2 Stutter""]/stutterVariance").Text)
'            End If
'        End If
        

    End With
    
End Sub


Function CheckAgainstDefault() As Boolean
'Check settings in cDeconSettings against Defaults and return boolean

    Dim tempCheck As Boolean: tempCheck = True

    If DefaultDeconSettings Is Nothing Then Call Factory.CreateDefaultSettings
    
    Dim Default As cDeconSettings: Set Default = DefaultDeconSettings
    
    Dim Checks As New Scripting.Dictionary
    'I really wish there were a way to iterate through object properties automatically, but there isn't. So:
    
    'Do each check and add its result to Checks dictionary
    Checks.Add "BurnAccepts", IIf(Me.BurnAccepts = Default.BurnAccepts, True, False)
    Checks.Add "PostBurnAccepts", IIf(Me.PostBurnAccepts = Default.PostBurnAccepts, True, False)
    Checks.Add "Chains", IIf(Me.Chains = Default.Chains, True, False)
    Checks.Add "RWSD", IIf(Me.RWSD = Default.RWSD, True, False)
    Checks.Add "PostBurnShortList", IIf(Me.PostBurnShortlist = Default.PostBurnShortlist, True, False)
    Checks.Add "MxPriors", IIf(Me.MxPriors = Default.MxPriors, True, False)
    Checks.Add "AutoContinueGR", IIf(Me.AutoContinueGR = Default.AutoContinueGR, True, False)
    Checks.Add "GRthreshold", IIf(Me.GRthreshold = Default.GRthreshold, True, False)
    Checks.Add "ExtraAccepts", IIf(Me.ExtraAccepts = Default.ExtraAccepts, True, False)
    
    Checks.Add "KitName", IIf(Me.KitName = Default.KitName, True, False)
    Checks.Add "AutosomalLoci", IIf(Me.AutosomalLoci = Default.AutosomalLoci, True, False)
    Checks.Add "Saturation", IIf(Me.Saturation = Default.Saturation, True, False)
    Checks.Add "DegradStart", IIf(Me.DegradStart = Default.DegradStart, True, False)
    Checks.Add "DegradMax", IIf(Me.DegradMax = Default.DegradMax, True, False)
    Checks.Add "DropInCap", IIf(Me.DropInCap = Default.DropInCap, True, False)
    Checks.Add "DropInFreq", IIf(Me.DropInFreq = Default.DropInFreq, True, False)
    Checks.Add "DropInGamma", IIf(Me.DropInGamma = Default.DropInGamma, True, False)
    Checks.Add "AlleleVariance", IIf(Me.AlleleVariance = Default.AlleleVariance, True, False)
    Checks.Add "MinVarianceFactor", IIf(Me.MinVarianceFactor = Default.MinVarianceFactor, True, False)
    Checks.Add "LocusAmpVariance", IIf(Me.LocusAmpVariance = Default.LocusAmpVariance, True, False)
    Checks.Add "VarianceMinParameter", IIf(Me.VarMinParameter = Default.VarMinParameter, True, False)
    
    'Check Stutter
    Dim v As Variant
    For Each v In Me.StutterSettings.Items
        Checks.Add v.StutterName & "_Variance", IIf(v.StutterVariance = Default.StutterSettings(v.StutterName).StutterVariance, True, False)
        Checks.Add v.StutterName & "_Max", IIf(v.StutterMax = Default.StutterSettings(v.StutterName).StutterMax, True, False)
    Next v
    
'    Checks.Add "BackStutterMax", IIf(Me.BackStutterMax = Default.BackStutterMax, True, False)
'    Checks.Add "BackStutterVariance", IIf(Me.BackStutterVariance = Default.BackStutterVariance, True, False)
'    Checks.Add "ForwardStutterMax", IIf(Me.ForwardStutterMax = Default.ForwardStutterMax, True, False)
'    Checks.Add "ForwardStutterVariance", IIf(Me.ForwardStutterVariance = Default.ForwardStutterVariance, True, False)
'    Checks.Add "SE33StutterMax", IIf(Me.SE33StutterMax = Default.SE33StutterMax, True, False)
'    Checks.Add "SE33StutterVariance", IIf(Me.SE33StutterVariance = Default.SE33StutterVariance, True, False)
    
    'VarNOC settings
    If Me.PopForRange <> "" Then Checks.Add "VarNOC_population", IIf(Me.PopForRange = Default.PopForRange, True, False)
    If Me.HRpercentAccepts <> 0 Then Checks.Add "VarNOC_HRaccepts", IIf(Me.HRpercentAccepts = Default.HRpercentAccepts, True, False)

   
    For Each v In Checks.Keys
        If Checks(v) = False Then
            tempCheck = False
            Debug.Print v & ": " & Checks(v)
        End If
    Next v
    
    CheckAgainstDefault = tempCheck

End Function




'****************************************************
'             cDeconSettings Properties
'****************************************************

Public Property Get STRmixVersion() As String
    STRmixVersion = pSTRmixVersion
End Property
Public Property Let STRmixVersion(v As String)
    pSTRmixVersion = v
End Property

Public Property Get TimeStampDecon() As String
    TimeStampDecon = pTimeStampDecon
End Property
Public Property Let TimeStampDecon(v As String)
    pTimeStampDecon = v
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

Public Property Get NOC() As Integer
    NOC = pNOC
End Property
Public Property Let NOC(v As Integer)
    pNOC = v
End Property


'MCMC Settings

Public Property Get BurnAccepts() As Long
    BurnAccepts = pBurnAccepts
End Property
Public Property Let BurnAccepts(v As Long)
    pBurnAccepts = v
End Property

Public Property Get PostBurnAccepts() As Long
    PostBurnAccepts = pPostBurnAccepts
End Property
Public Property Let PostBurnAccepts(v As Long)
    pPostBurnAccepts = v
End Property

Public Property Get Chains() As Integer
    Chains = pChains
End Property
Public Property Let Chains(v As Integer)
    pChains = v
End Property

Public Property Get RWSD() As Double
    RWSD = pRWSD
End Property
Public Property Let RWSD(v As Double)
    pRWSD = v
End Property

Public Property Get PostBurnShortlist() As Integer
    PostBurnShortlist = pPostBurnShortlist
End Property
Public Property Let PostBurnShortlist(v As Integer)
    pPostBurnShortlist = v
End Property

Public Property Get MxPriors() As Boolean
    MxPriors = pMxPriors
End Property
Public Property Let MxPriors(v As Boolean)
    pMxPriors = v
End Property

Public Property Get AutoContinueGR() As Boolean
    AutoContinueGR = pAutoContinueGR
End Property
Public Property Let AutoContinueGR(v As Boolean)
    pAutoContinueGR = v
End Property

Public Property Get GRthreshold() As Double
    GRthreshold = pGRthreshold
End Property
Public Property Let GRthreshold(v As Double)
    pGRthreshold = v
End Property

Public Property Get ExtraAccepts() As Long
    ExtraAccepts = pExtraAccepts
End Property
Public Property Let ExtraAccepts(v As Long)
    pExtraAccepts = v
End Property

Public Property Get HRpercentAccepts() As Double
    HRpercentAccepts = pHRpercentAccepts
End Property
Public Property Let HRpercentAccepts(v As Double)
    pHRpercentAccepts = v
End Property

Public Property Get PopForRange() As String
    PopForRange = pPopForRange
End Property
Public Property Let PopForRange(v As String)
    pPopForRange = v
End Property


'Kit Settings

Public Property Get KitName() As String
    KitName = pKitName
End Property
Public Property Let KitName(v As String)
    pKitName = v
End Property

Public Property Get AutosomalLoci() As Integer
    AutosomalLoci = pAutosomalLoci
End Property
Public Property Let AutosomalLoci(v As Integer)
    pAutosomalLoci = v
End Property

Public Property Get Saturation() As Long
    Saturation = pSaturation
End Property
Public Property Let Saturation(v As Long)
    pSaturation = v
End Property

Public Property Get DegradStart() As Double
    DegradStart = pDegradStart
End Property
Public Property Let DegradStart(v As Double)
    pDegradStart = v
End Property

Public Property Get DegradMax() As Double
    DegradMax = pDegradMax
End Property
Public Property Let DegradMax(v As Double)
    pDegradMax = v
End Property

Public Property Get DropInCap() As Integer
    DropInCap = pDropInCap
End Property
Public Property Let DropInCap(v As Integer)
    pDropInCap = v
End Property

Public Property Get DropInFreq() As Double
    DropInFreq = pDropInFreq
End Property
Public Property Let DropInFreq(v As Double)
    pDropInFreq = v
End Property

Public Property Get DropInGamma() As String
    DropInGamma = pDropInGamma
End Property
Public Property Let DropInGamma(v As String)
    pDropInGamma = v
End Property

Public Property Get AlleleVariance() As String
    AlleleVariance = pAlleleVariance
End Property
Public Property Let AlleleVariance(v As String)
    pAlleleVariance = v
End Property

Public Property Get MinVarianceFactor() As Double
    MinVarianceFactor = pMinVarianceFactor
End Property
Public Property Let MinVarianceFactor(v As Double)
    pMinVarianceFactor = v
End Property

Public Property Get LocusAmpVariance() As Double
    LocusAmpVariance = pLocusAmpVariance
End Property
Public Property Let LocusAmpVariance(v As Double)
    pLocusAmpVariance = v
End Property

Public Property Get VarMinParameter() As Double
    VarMinParameter = pVarMinParameter
End Property
Public Property Let VarMinParameter(v As Double)
    pVarMinParameter = v
End Property



Public Property Get StutterSettings() As Dictionary
    Set StutterSettings = pStutterSettings
End Property
Public Property Let StutterSettings(v As Dictionary)
    Set pStutterSettings = v
End Property


'Public Property Get BackStutterMax() As Double
'    BackStutterMax = pBackStutterMax
'End Property
'Public Property Let BackStutterMax(v As Double)
'    pBackStutterMax = v
'End Property
'
'Public Property Get ForwardStutterMax() As Double
'    ForwardStutterMax = pForwardStutterMax
'End Property
'Public Property Let ForwardStutterMax(v As Double)
'    pForwardStutterMax = v
'End Property
'
'Public Property Get SE33StutterMax() As Double
'    SE33StutterMax = pSE33StutterMax
'End Property
'Public Property Let SE33StutterMax(v As Double)
'    pSE33StutterMax = v
'End Property
'
'Public Property Get BackStutterVariance() As String
'    BackStutterVariance = pBackStutterVariance
'End Property
'Public Property Let BackStutterVariance(v As String)
'    pBackStutterVariance = v
'End Property
'
'Public Property Get ForwardStutterVariance() As String
'    ForwardStutterVariance = pForwardStutterVariance
'End Property
'Public Property Let ForwardStutterVariance(v As String)
'    pForwardStutterVariance = v
'End Property
'
'Public Property Get SE33StutterVariance() As String
'    SE33StutterVariance = pSE33StutterVariance
'End Property
'Public Property Let SE33StutterVariance(v As String)
'    pSE33StutterVariance = v
'End Property

Public Property Get AutoSearch() As Boolean
    AutoSearch = pAutoSearch
End Property
Public Property Let AutoSearch(v As Boolean)
    pAutoSearch = v
End Property



