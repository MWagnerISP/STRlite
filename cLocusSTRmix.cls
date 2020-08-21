VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLocusSTRmix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Represents a locus deconvoluted by STRmix
'Here there are multiple possible genotype pairings, each stored in AllGenotypes as cGenotypeSTRmix objects
'(whereas with the cLocusWhole, the genotype is one set of all alleles at the locus and doesn't need distinct objects)
'This is where the deconvoluted locus is analyzed for CODIS suitability and stored in PulledGenotypes if it passes
'The FinalPull is the summary of the genotype pairings that, collectively, are CODIS eligible
'***************************************************************************************************

Private pLocusName As String
Private pCore13 As Boolean 'whether the locus is a core 13 CODIS locus
Private pCore20 As Boolean 'whether the locus is a core 20 CODIS locus
'Private pAutosomal As Boolean 'whether the locus is autosomal or not

Private pAllGenotypes As New Scripting.Dictionary 'key = Genotype string, item = cGenotypeSTRmix object. All genotypes in descending order of weight %

Private pPulledGenotypes As New Scripting.Dictionary 'key = Genotype string, item = cGenotypeSTRmix object
Private pPulledAlleles As New Scripting.Dictionary
Private pLocusPulled As Boolean 'whether or not the locus passed the test
Private pObligateAllele As String
Private pFinalPull As String 'end result of test for this locus; what goes on the Decon tab
Private pLocusCategory As String
Private pMME As Double 'moderate match estimate
Private pMRE As Double 'match rarity estimate

Dim gen As Variant, allele As Variant

'The following dictionaries don't need to be saved as properties
Dim DropoutGenotypes As Scripting.Dictionary 'key = obligate allele, item = cGenotypeSTRmix object
Dim ThresholdGenotypes As Scripting.Dictionary 'key = Genotype string, item = cGenotypeSTRmix object
Dim ThresholdAlleles As Scripting.Dictionary 'key = allele, item = obligate T/F

Option Explicit


Sub InitializeMe(LocusName As String, GenotypesRange As Range)

    Dim rng As Range
    Dim newGenotype As cGenotypeSTRmix
    Dim tempDict As New Scripting.Dictionary
    
    If CODISThreshold = 0 Or CODISMaxAlleles = 0 Then Call Admin.LoadSTRliteSettings
    If CODISCore13 Is Nothing Or CODISCore20 Is Nothing Then Call Admin.LociCollection
    
    With Me
        .LocusName = LocusName
        .Core13 = CODISCore13.Exists(Me.LocusName)
        .Core20 = CODISCore20.Exists(Me.LocusName)
        .LocusPulled = False 'assume locus doesn't pass. It has to earn it!
    End With
    
    'Make a new cGenotypeSTRmix object for each genotype and add to temp dictionary
    For Each rng In GenotypesRange
        Set newGenotype = Factory.CreateGenotype(rng, Me.LocusName) 'pass the current locus name to the genotype
        tempDict.Add newGenotype.Genotype, newGenotype
    Next rng
    
    'Make the AllGenotypes dictionary sorted by descending genotype weight
    Me.AllGenotypes = SortObjectsByProperty(tempDict, "Percent", xlDescending)
    
    'Check for total dropout
    If AllGenotypes.Exists("-1, -1") Then
        If AllGenotypes("-1, -1").Percent > (1 - CODISThreshold) Then
            Me.LocusCategory = "None"
            Me.LocusPulled = False
            Me.FinalPull = ""
            Exit Sub
        End If
    End If
    
    'Check for single source
    If AllGenotypes.Items()(0).Percent >= CODISThreshold Then 'The first Genotype in AllGenotypes will have the highest weight
        For Each allele In AllGenotypes.Items()(0).Alleles
            If Not Me.PulledAlleles.Exists(CDbl(allele)) Then Me.PulledAlleles.Add CDbl(allele), CDbl(allele)
        Next allele
        Me.LocusCategory = "U"
        Me.LocusPulled = True
        Me.FinalPull = AllGenotypes.Items()(0).Genotype
        Me.PulledGenotypes.Add AllGenotypes.Items()(0).Genotype, AllGenotypes.Items()(0)
        Me.MRE = Me.PulledGenotypes.Items()(0).RMP 'MRE = high stringency locus = same as RMP
        If PulledAlleles.Count = 1 Then
            Me.ObligateAllele = PulledAlleles.Items()(0)
            Me.MME = CalculateMME(PulledAlleles, "ObligateOnly")
        Else:
            Me.MME = CalculateMME(PulledAlleles, "Mixture")
        End If
        Exit Sub
    End If
    
    Call Me.AnalyzeGenotypes

    Application.StatusBar = False
    
End Sub


Sub AnalyzeGenotypes()
'At this point, only mixtures and partials because we've weeded out Single Sources

    Application.StatusBar = "Analyzing " & Me.LocusName

    Set ThresholdGenotypes = New Scripting.Dictionary
    Set ThresholdAlleles = New Scripting.Dictionary
    Set DropoutGenotypes = New Scripting.Dictionary
    
    Dim MaxAlleles As Integer: MaxAlleles = CODISMaxAlleles
    Dim weight As Double: weight = 0

    '-------------------------------------------
    '  TRY TO REACH THRESHOLD WITHOUT DROPOUT
    '-------------------------------------------
    For Each gen In Me.AllGenotypes.Items
    
        Select Case gen.Alleles.Exists("-1") 'Check if genotype has a "-1" allele (dropout)
        
            Case False  'The genotype does not have dropout.
                'Add the genotype to ThresholdGenotypes until weight reaches threshold.
                
                If CDec(weight) <= CDec(CODISThreshold) Then 'Only add genotype if threshold hasn't been reached
                'Add genotype percent to running weight total.
                    weight = weight + gen.Percent
                    
                'Add genotype to ThresholdGenotypes list
                    ThresholdGenotypes.Add gen.Genotype, gen
                    
                'Add alleles to ThresholdAlleles list
                    For Each allele In gen.Alleles.Keys
                        If Not ThresholdAlleles.Exists(allele) Then ThresholdAlleles.Add allele, True
                        'Start with considering all alleles obligate (True)
                    Next allele
                End If
                
            
            Case True  'The genotype contains dropout.
                'Save the genotype in DropoutGenotypes in case we need it later.
                Dim Obligate As String
                If gen.Genotype <> "-1" Then 'don't count total dropout genotypes.
                    For Each allele In gen.Alleles.Keys
                        If allele <> "-1" Then Obligate = allele 'Save the non-dropout allele as a possible obligate
                    Next allele
                    If Not DropoutGenotypes.Exists(Obligate) Then DropoutGenotypes.Add Obligate, gen
                End If
            
        End Select
    
    Next gen
    
    '-------------------------------------------
    '  CHECK STATUS
    '-------------------------------------------
    
    If CDec(weight) >= CDec(CODISThreshold) Then
    
        '-------------------------------------------
        '  THRESHOLD ACHIEVED!
        '-------------------------------------------
        'Check for obligates
        If Me.CheckForObligates = True Then
            Me.LocusPulled = True
            'Check for allele count
            If ThresholdAlleles.Count > MaxAlleles Then
                Me.Finalize ("ObligateOnly") 'ObligateOnly= Too many total alleles, but obligate exists
            Else: Me.Finalize ("MixWithObligate") 'MixWithObligate = Fewer than max total alleles, including obligate
            End If
            
        Else:
            
            If ThresholdAlleles.Count <= MaxAlleles Then
                Me.Finalize ("Mixture") 'Mixture = Fewer than max total alleles, no obligate
            Else:
            
                'If there are >MaxAlleles and no obligate, try holding each allele obligate to see if one works
                'Note: ThresholdAlleles may be missing alleles from low-weighted genotypes, but any missing alleles won't be obligate by definition.
                For Each allele In ThresholdAlleles.Keys
                    If TestThreshold(allele) Then
                        Me.ObligateAllele = allele
                        Me.LocusPulled = True
                        Me.Finalize ("ObligateOnly")
                        Exit For
                    End If
                Next allele
                'If that didn't work, then game over
                If Me.LocusPulled = False Then Me.Finalize ("Nope")
            End If
        End If
        
        '-------------------------------------------
        '  THRESHOLD NOT ACHIEVED...
        '-------------------------------------------
        'Didn't reach weight threshold. Try holding DropoutGenotype alleles obligate to see if one works.
    Else:
            
            For Each allele In DropoutGenotypes.Keys 'the keys are obligate alleles
                If TestThreshold(allele) Then
                    Me.LocusPulled = True
                    Me.ObligateAllele = allele
                    Me.Finalize ("ObligateOnly")
                    Exit For
                End If
            Next allele
            'If that didn't work, then game over
            If Me.LocusPulled = False Then Me.Finalize ("Nope")
        
    End If
    
End Sub

Function TestThreshold(Obligate As Variant) As Boolean
'Tests genotype weights against the threshold by holding one allele obligate
'Returns True if the threshold can be met using only genotypes with the obligate allele

    Dim weight As Double: weight = 0
    
    ThresholdGenotypes.RemoveAll 'Clear the ThresholdGenotypes dictionary
    
    For Each gen In Me.AllGenotypes.Items
        If gen.Alleles.Exists(Obligate) Then 'Checks if the genotype contains the obligate allele
            'Only adds the weight of genotypes with the obligate allele
            weight = weight + gen.Percent
            ThresholdGenotypes.Add gen.Genotype, gen
        End If
    Next gen
        
    TestThreshold = (weight >= CDec(CODISThreshold))
    
End Function

Function CheckForObligates() As Boolean
'Returns True if there is an obligate allele, and saves that allele as ObligateAllele property

    'Check each allele in ThresholdAlleles against each genotype
    For Each allele In ThresholdAlleles.Keys
        For Each gen In ThresholdGenotypes.Items 'gen is a cGenotypeSTRmix object
            'All alleles start as "True" and get turned off if they're not obligate
            'If there is a genotype that does not contain the allele, the allele is not obligate
            If Not gen.Alleles.Exists(allele) Then ThresholdAlleles(allele) = False
        Next gen
    Next allele
    
    'Check if there are any obligates remaining & save as ObligateAllele
    For Each allele In ThresholdAlleles.Keys
        If ThresholdAlleles(allele) = True Then
            Me.ObligateAllele = allele
            CheckForObligates = True
            Exit For
        End If
    Next allele
    
End Function


Sub Finalize(ResultType As String)
'Consolidates the pulled genotypes into a string
'Marks obligate alleles & multiple genotype loci
'Designates locus category
'Calculates MME

    'Dim PulledAlleles As New Scripting.Dictionary

    'Save ThresholdGenotypes into PulledGenotypes property
    For Each gen In ThresholdGenotypes.Items
        PulledGenotypes.Add gen.Genotype, gen
        For Each allele In gen.Alleles
            'Make alleles Double type so that they'll sort in numerical order
            If Not PulledAlleles.Exists(CDbl(allele)) Then PulledAlleles.Add CDbl(allele), CDbl(allele)
        Next allele
    Next gen

    Select Case ResultType
    
        Case "ObligateOnly"
            Me.LocusPulled = True
            Me.LocusCategory = "P"
            Me.FinalPull = Me.ObligateAllele & "+"
            
        Case "Mixture"
            Me.LocusPulled = True
            Me.LocusCategory = "M"
            Me.FinalPull = CreateFinalPull(PubFun.SortDictionaryByKey(PulledAlleles, xlAscending))
        
        Case "MixWithObligate"
            Me.LocusPulled = True
            Me.LocusCategory = "M"
            Me.FinalPull = CreateFinalPull(PubFun.SortDictionaryByKey(PulledAlleles, xlAscending), Me.ObligateAllele)
        
        Case "Nope"
            Me.LocusPulled = False
            Me.LocusCategory = "None"
            Me.FinalPull = ""
    
    End Select
    
    If ResultType <> "Nope" Then Me.MME = CalculateMME(PulledAlleles, ResultType)

End Sub

Function CalculateMME(PulledAlleles As Dictionary, ResultType As String) As Double
'Match Estimation Calculations taken from "CODIS Online Help" version 8.0
'There are three different equations to calculate MME for the three ResultTypes we've already determined:

    Dim freqDict As Scripting.Dictionary: Set freqDict = New Scripting.Dictionary
    Dim tempDict As Dictionary: Set tempDict = LociSTRmix(Me.LocusName) 'shortcut for locus allele frequency dictionary
    
    'Make a more manageable dictionary with the alleles in question and their frequencies
    For Each allele In PulledAlleles.Keys
        If tempDict.Exists(CDbl(allele)) Then
            freqDict.Add CStr(allele), tempDict(CDbl(allele))
        Else: 'if allele isn't listed, use the min allele freq (except for -1 allele)
            'If allele <> -1 Then
                'Debug.Print Me.LocusName & "_Min Allele Freq: " & allele
                freqDict.Add CStr(allele), tempDict("min")
            'End If
        End If
    Next allele
    
    Dim obl As Double 'frequency of the obligate allele
    Dim nonobl As Double 'SUM of frequencies of non-obligate alleles

    Select Case ResultType
    
        Case "ObligateOnly"
            obl = freqDict(CStr(Me.ObligateAllele))
            CalculateMME = (2 * obl) - (obl * obl)
            
        Case "Mixture"
            nonobl = 0
            For Each allele In freqDict.Keys
                nonobl = nonobl + freqDict(CStr(allele))
            Next allele
            
            CalculateMME = nonobl * nonobl
        
        Case "MixWithObligate"
            nonobl = 0
            For Each allele In freqDict.Keys
                If allele <> Me.ObligateAllele Then nonobl = nonobl + freqDict(CStr(allele))
            Next allele
        
            obl = freqDict(CStr(Me.ObligateAllele))
            CalculateMME = (obl * obl) + (2 * obl * nonobl)
    
    End Select

    'Debug.Print Me.LocusName & " MME: " & CalculateMME

End Function


Function CreateFinalPull(alleledict As Dictionary, Optional Obligate As String = "") As String
    'AlleleDict has been sorted by ascending key
    'Only used for mixtures
    
    'If there's an obligate, mark it with "+"
    If Obligate <> "" Then alleledict(CDbl(Obligate)) = Obligate & "+"
    
    'Combine all alleles into one string
    Dim AlleleSet() As Variant
    AlleleSet = alleledict.Items()
    CreateFinalPull = Join(AlleleSet, ", ")
    
    'Put braces around FinalPull if it's ambiguous
    If alleledict.Count <= 2 And Obligate = "" And Me.PulledGenotypes.Count > 1 Then
        CreateFinalPull = "{" & CreateFinalPull & "}"
    End If

End Function


'****************************************************
'                cLocusSTRmix Properties
'****************************************************

Public Property Get LocusName() As String
    LocusName = pLocusName
End Property
Public Property Let LocusName(v As String)
    pLocusName = v
End Property

Public Property Get Core13() As Boolean
    Core13 = pCore13
End Property
Public Property Let Core13(v As Boolean)
    pCore13 = v
End Property

Public Property Get Core20() As Boolean
    Core20 = pCore20
End Property
Public Property Let Core20(v As Boolean)
    pCore20 = v
End Property

'Public Property Get Autosomal() As Boolean
'    Autosomal = pAutosomal
'End Property
'Public Property Let Autosomal(v As Boolean)
'    pAutosomal = v
'End Property


Public Property Get AllGenotypes() As Dictionary
    Set AllGenotypes = pAllGenotypes
End Property
Public Property Let AllGenotypes(v As Dictionary)
    Set pAllGenotypes = v
End Property

Public Property Get PulledGenotypes() As Dictionary
    Set PulledGenotypes = pPulledGenotypes
End Property
Public Property Let PulledGenotypes(v As Dictionary)
    Set pPulledGenotypes = v
End Property

Public Property Get PulledAlleles() As Dictionary
    Set PulledAlleles = pPulledAlleles
End Property
Public Property Let PulledAlleles(v As Dictionary)
    Set pPulledAlleles = v
End Property

Public Property Get LocusPulled() As Boolean
    LocusPulled = pLocusPulled
End Property
Public Property Let LocusPulled(v As Boolean)
    pLocusPulled = v
End Property

Public Property Get ObligateAllele() As String
    ObligateAllele = pObligateAllele
End Property
Public Property Let ObligateAllele(v As String)
    pObligateAllele = v
End Property

Public Property Get FinalPull() As String
    FinalPull = pFinalPull
End Property
Public Property Let FinalPull(v As String)
    pFinalPull = v
End Property

Public Property Get LocusCategory() As String
    LocusCategory = pLocusCategory
End Property
Public Property Let LocusCategory(v As String)
    pLocusCategory = v
End Property

Public Property Get MME() As Double
    MME = pMME
End Property
Public Property Let MME(v As Double)
    pMME = v
End Property

Public Property Get MRE() As Double
    MRE = pMRE
End Property
Public Property Let MRE(v As Double)
    pMRE = v
End Property
