VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cContributor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Represents a STRmix-determined contributor from a deconvolution.
'Contains all the genotype data per locus (stored in LociAll) and the "pulled" loci in LociPulled
'***************************************************************************************************

Private pContributorName As String
Private pSpecimenCategory As String
Private pConditioned As Boolean
Private pLociAll As New Scripting.Dictionary 'key = locus name, item = cLocusSTRmix object
Private pLociPulled As New Scripting.Dictionary 'key = locus name, item = cLocusSTRmix object

Private pMME As Double 'core 13 loci, moderate stringency + unambiguous RMP
Private pSDIS_MRE As Double 'any locus, high stringency only
Private pNDIS_MRE As Double 'core 13 loci, high stringency only

Dim Locus As Variant
Dim loc As Variant, rng As Range, i As Integer


Public Sub InitializeMe(DeconID As String, ContributorNum As Integer)

    Dim FullList As Range
'    Set FullList = Sheets("Import").Range("A1:A" & Cells(Rows.Count, "A").End(xlUp).Row)  'This one was unreliable?
    Set FullList = Sheets("Import").Range("A1:A" & Sheets("Import").UsedRange.Rows.Count)
    Dim LocusHeader As Range
    Dim ContributorHeader As Range
    Dim GenotypeRange As Range
    Dim newLocus As cLocusSTRmix
    
    Me.ContributorName = "Contributor " & ContributorNum 'will be replaced later if conditioned
    
    If LociSTRmix Is Nothing Or LociSDIS Is Nothing Then Call Admin.LociCollection
    
    'Debug.Print Me.ContributorName
    
    For Each loc In LociSTRmix.Keys
        'Find headers for locus & contributor (within locus)
            Set LocusHeader = PubFun.FindCell(FullList, "Locus " & loc, Range("A1"))
            Set ContributorHeader = PubFun.FindCell(FullList, "Contributor " & ContributorNum, LocusHeader)
        
        'Get range of genotypes for this locus/contributor
            i = 1
            Set rng = ContributorHeader.Offset(i, 0) 'start with first cell under the Contributor heading
            Set GenotypeRange = rng 'GenotypeRange can't start out as Nothing so set it = rng
            If InStr(1, GenotypeRange.Value, "Genotype") = 0 Then GoTo NoLocus
            
            Do Until InStr(1, rng.Value, "Genotype") = 0 'Do this until rng doesn't contain the word "Genotype"
                Set GenotypeRange = Union(GenotypeRange, rng)
                i = i + 1
                Set rng = ContributorHeader.Offset(i, 0) 'set rng to next row after header
            Loop
            
        'Create cLocusSTRmix object & add it to LociAll
            Set newLocus = Factory.CreateLocusSTRmix(CStr(loc), GenotypeRange)
            If Not Me.LociAll.Exists(newLocus.LocusName) Then Me.LociAll.Add newLocus.LocusName, newLocus
            
NoLocus:

    Next loc
    
    'Detect pulled loci, specimen category, & MME/MRE
    Me.ProfileResults

End Sub

Sub ProfileResults()
'Figures out which loci were able to be pulled for CODIS and which categories

    Dim UnkCat As Integer, MixCat As Integer, PartCat As Integer, CODIS13Count As Integer
    Dim Pulled_Core13 As New Dictionary
    Dim Pulled_SDIS As New Dictionary
    
    'Count how many loci were pulled, how many of each locus category were pulled, and how many core CODIS loci were pulled
        For Each Locus In LociAll.Items
            If Locus.LocusPulled Then
            
                LociPulled.Add Locus.LocusName, Locus
                
                If Locus.LocusCategory = "U" Then UnkCat = UnkCat + 1

                If Locus.Core13 Then
                    CODIS13Count = CODIS13Count + 1
                    Pulled_Core13.Add Locus.LocusName, Locus
                End If
                
                'If locus is one of the loci selected for SDIS MRE, add to Pulled_SDIS
                If LociSDIS.Exists(Locus.LocusName) Then Pulled_SDIS.Add Locus.LocusName, Locus
                
                If Locus.LocusCategory = "M" Then MixCat = MixCat + 1
                If Locus.LocusCategory = "P" Then PartCat = PartCat + 1
                
            End If
        Next Locus
    
    'Assign specimen categories
        If Me.LociPulled.Count = UnkCat Then Me.SpecimenCategory = "Unknown" 'If all loci pulled are "U", then SpecimenCategory = "Unknown"
        If PartCat > 0 Or CODIS13Count < 13 Then Me.SpecimenCategory = "Partial" 'If there are any partial loci OR if there are any Core 13 CODIS loci missing, partial trumps unknown
        If MixCat > 0 Then Me.SpecimenCategory = "Mixture" 'If there are any mixture loci, mixture trumps partial


    'Calculate MME/MRE
        
    Me.MME = CalculateMME(Pulled_Core13)
    Me.SDIS_MRE = CalculateMRE(Pulled_SDIS) 'SDIS considers loci selected in STRlite Settings
    Me.NDIS_MRE = CalculateMRE(Pulled_Core13) 'NDIS MRE considers Core13 only
    
End Sub

Function CalculateMME(LociSource As Dictionary) As Double
'MME is only calculated on Core13. LociSource is a dictionary that contains only core13 loci
'Each locus already has a locus-level MME calculated

    Dim tempMME As Double: tempMME = 1
    
    For Each Locus In LociSource.Items
        If Locus.MME > 0 Then tempMME = tempMME * Locus.MME
        'Debug.Print Locus.LocusName & " MME: " & Locus.MME
    Next Locus
    
    CalculateMME = tempMME

End Function

Function CalculateMRE(LociSource As Dictionary) As Double
'LociSource is a dictionary that contains the loci to calculate
'Each single-genotype locus already has a locus-level MRE calculated

    Dim tempMRE As Double: tempMRE = 1
    
    For Each Locus In LociSource.Items
        'Debug.Print Locus.LocusName & " MRE: " & Locus.MRE
        Select Case Locus.LocusCategory
            Case "U"
                tempMRE = tempMRE * Locus.MRE
            Case Else
                tempMRE = tempMRE * Locus.MME
        End Select
        
        'Debug.Print Locus.MRE
    Next Locus
    
    CalculateMRE = tempMRE
    
End Function


'****************************************************
'            cContributor Properties
'****************************************************

Public Property Get ContributorName() As String
    ContributorName = pContributorName
End Property
Public Property Let ContributorName(v As String)
    pContributorName = v
End Property

Public Property Get SpecimenCategory() As String
    SpecimenCategory = pSpecimenCategory
End Property
Public Property Let SpecimenCategory(v As String)
    pSpecimenCategory = v
End Property

Public Property Get Conditioned() As Boolean
    Conditioned = pConditioned
End Property
Public Property Let Conditioned(v As Boolean)
    pConditioned = v
End Property

Public Property Get LociAll() As Dictionary
    Set LociAll = pLociAll
End Property
Public Property Let LociAll(v As Dictionary)
    Set pLociAll = v
End Property

Public Property Get LociPulled() As Dictionary
    Set LociPulled = pLociPulled
End Property
Public Property Let LociPulled(v As Dictionary)
    Set pLociPulled = v
End Property



Public Property Get MME() As Double
    MME = pMME
End Property
Public Property Let MME(v As Double)
    pMME = v
End Property

Public Property Get SDIS_MRE() As Double
    SDIS_MRE = pSDIS_MRE
End Property
Public Property Let SDIS_MRE(v As Double)
    pSDIS_MRE = v
End Property

Public Property Get NDIS_MRE() As Double
    NDIS_MRE = pNDIS_MRE
End Property
Public Property Let NDIS_MRE(v As Double)
    pNDIS_MRE = v
End Property
