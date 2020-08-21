VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLRSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pTimeStampDecon As String

Private pCaseNum As String
Private pSampleID As String

Private pAssignSubSourceLR As Boolean
Private pCalculateHPD As Boolean
Private pHPDiterations As Long
Private pMCMCuncertainty As Boolean
Private pAlleleFreqUncertainty As Boolean
Private pHPDquantile As Integer
Private pHPDsides As Integer

Public Sub InitializeMe(rootConfig As Object, TimeStampID As String)

'rootConfig = IXMLDOMElement (root element of LRConfigXML)
'the TimeStampID is to label this particular cLRSettings object. It comes from the sister LR file, Results.xml

    'Harvest all the config settings from rootConfig. Most are single unique nodes (tags).
    With Me
        
        .TimeStampDecon = TimeStampID
        .CaseNum = CStr(rootConfig.selectSingleNode("//caseNumber").Text)
        .SampleID = CStr(rootConfig.selectSingleNode("//sampleID").Text) 'note: "sampleId" in results.xml but "sampleID" in config.xml
    
        'v2.5 has "displayFactorOfN" instead of "assignSubSourceLr"
        If rootConfig.selectSingleNode("//assignSubSourceLr") Is Nothing Then
            .AssignSubSourceLR = CBool(rootConfig.selectSingleNode("//displayFactorOfN").Text)
        Else: .AssignSubSourceLR = CBool(rootConfig.selectSingleNode("//assignSubSourceLr").Text)
        End If
        
        .CalculateHPD = CBool(rootConfig.selectSingleNode("//calculateHpd").Text)
        
        If .CalculateHPD Then
            .HPDiterations = CLng(rootConfig.selectSingleNode("//hpdIterations").Text)
            .MCMCuncertainty = CBool(rootConfig.selectSingleNode("//includeMcmcUncertainty").Text)
            If rootConfig.selectSingleNode("//includeAlleleFreqUncertainty") Is Nothing Then 'v2.5 doesn't have this
                .AlleleFreqUncertainty = False
                Else: .AlleleFreqUncertainty = CBool(rootConfig.selectSingleNode("//includeAlleleFreqUncertainty").Text)
            End If
            .HPDquantile = CDbl(rootConfig.selectSingleNode("//quantile").Text)
            .HPDsides = CInt(rootConfig.selectSingleNode("//sides").Text)
        End If

 
    End With

End Sub

Function CheckAgainstDefault() As Boolean

'Check settings in cLRSettings against Defaults and return boolean

    Dim tempCheck As Boolean: tempCheck = True

    If DefaultLRSettings Is Nothing Then Call Factory.CreateDefaultLRSettings
    
    Dim Checks As New Scripting.Dictionary
    
    'Do each check and add its result to Checks dictionary
    Checks.Add "SubSource LR", IIf(Me.AssignSubSourceLR = DefaultLRSettings.AssignSubSourceLR, True, False)
    Checks.Add "Calculate HPD", IIf(Me.CalculateHPD = DefaultLRSettings.CalculateHPD, True, False)
    Checks.Add "HPD Iterations", IIf(Me.HPDiterations = DefaultLRSettings.HPDiterations, True, False)
    Checks.Add "MCMC Uncertainty", IIf(Me.MCMCuncertainty = DefaultLRSettings.MCMCuncertainty, True, False)
    Checks.Add "Allele Frequency Uncertainty", IIf(Me.AlleleFreqUncertainty = DefaultLRSettings.AlleleFreqUncertainty, True, False)
    Checks.Add "HPD quantile", IIf(Me.HPDquantile = DefaultLRSettings.HPDquantile, True, False)
    Checks.Add "HPD sides", IIf(Me.HPDsides = DefaultLRSettings.HPDsides, True, False)

    Dim v As Variant
    For Each v In Checks.Keys
        If Checks(v) = False Then
            tempCheck = False
            'Debug.Print v & ": " & Checks(v)
        End If
    Next v
    
    CheckAgainstDefault = tempCheck

End Function



Public Property Get TimeStampDecon() As String
    TimeStampDecon = pTimeStampDecon
End Property
Public Property Let TimeStampDecon(v As String)
    pTimeStampDecon = v
End Property

Public Property Get TimeStampLR() As String
    TimeStampLR = pTimeStampLR
End Property
Public Property Let TimeStampLR(v As String)
    pTimeStampLR = v
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

Public Property Get AssignSubSourceLR() As Boolean
    AssignSubSourceLR = pAssignSubSourceLR
End Property
Public Property Let AssignSubSourceLR(v As Boolean)
    pAssignSubSourceLR = v
End Property

Public Property Get CalculateHPD() As Boolean
    CalculateHPD = pCalculateHPD
End Property
Public Property Let CalculateHPD(v As Boolean)
    pCalculateHPD = v
End Property

Public Property Get HPDiterations() As Long
    HPDiterations = pHPDiterations
End Property
Public Property Let HPDiterations(v As Long)
    pHPDiterations = v
End Property

Public Property Get MCMCuncertainty() As Boolean
    MCMCuncertainty = pMCMCuncertainty
End Property
Public Property Let MCMCuncertainty(v As Boolean)
    pMCMCuncertainty = v
End Property

Public Property Get AlleleFreqUncertainty() As Boolean
    AlleleFreqUncertainty = pAlleleFreqUncertainty
End Property
Public Property Let AlleleFreqUncertainty(v As Boolean)
    pAlleleFreqUncertainty = v
End Property

Public Property Get HPDquantile() As Integer
    HPDquantile = pHPDquantile
End Property
Public Property Let HPDquantile(v As Integer)
    pHPDquantile = v
End Property

Public Property Get HPDsides() As Integer
    HPDsides = pHPDsides
End Property
Public Property Let HPDsides(v As Integer)
    pHPDsides = v
End Property

