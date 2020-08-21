VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGenotypeSTRmix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Represents a single allele pairing, per locus, per contributor, output from a STRmix decon.
'Includes the possible genotype and its weighting
'This genotype data comes from the ComponentInterp text file.
'***************************************************************************************************

Private pLocusName As String 'we only really need this to get the allele frequencies
Private pGenotype As String 'Formatted "Allele1, Allele2"
Private pPercent As Double 'genotype weighting
Private pRMP As Double 'actual RMP for this genotype. Used for high-stringency MRE
Private pAlleles As New Scripting.Dictionary 'Will make it easier to detect if a certain allele is present in the genotype


Sub InitializeMe(rngGen As Range, Locus As String)
'rngGen is the actual cell containing the genotype to be loaded

    Dim tempGen As String, tempPerc As String, arrAlleles As Variant

    'Grab the genotype between [ ]
        tempGen = LTrim(RTrim(Mid(rngGen.Value, 12, (InStr(1, rngGen.Value, "]") - 12))))
        'Debug.Print "Genotype: " & tempGen
        
    'Grab the percentage between the - and %
        tempPerc = LTrim(Replace(Right(rngGen.Value, Len(rngGen.Value) - InStr(1, rngGen.Value, "] - ") - 2), "%", ""))
        Me.Percent = CDbl(tempPerc) / 100
        'Debug.Print "Weight: " & Me.Percent
        
    'Split genotype into alleles & store in dictionary
        arrAlleles = Split(tempGen, ", ")
        Alleles.Add arrAlleles(0), arrAlleles(0)
        If Not Alleles.Exists(arrAlleles(1)) Then Alleles.Add arrAlleles(1), arrAlleles(1)
        
    'Calculate RMP
        Me.LocusName = Locus
        Me.RMP = CalculateRMP(Me.LocusName, CDbl(arrAlleles(0)), CDbl(arrAlleles(1)))
        
    'Recombine the allele(s) into the genotype. Consolidates homozygotes at this step
        Me.Genotype = Join(Alleles.Keys, ", ")
    
End Sub


Function CalculateRMP(Locus As String, A1 As Double, A2 As Double) As Double
'NRC '96 RMP calculation

Dim tempRMP As Double
Dim tempDict As Dictionary 'placeholder for the dictionary that contains this locus' allele frequencies
Set tempDict = LociSTRmix(Locus)
Dim p As Double, q As Double

'If the allele has a frequency, use it, else use the minimum allele frequency labeled "min"

If tempDict.Exists(A1) Then
    p = tempDict(A1)
Else:
    p = tempDict("min")
    'Debug.Print Me.LocusName & "_Min Allele Freq: " & A1
End If

If tempDict.Exists(A2) Then
    q = tempDict(A2)
Else:
    q = tempDict("min")
    'Debug.Print Me.LocusName & "_Min Allele Freq: " & A2
End If
    
'CANNOT use IIf statement because it adds alleles without frequencies into tempDict --> LociSTRmix
'p = IIf(tempDict.Exists(A1), tempDict(A1), tempDict("min"))
'q = IIf(tempDict.Exists(A2), tempDict(A2), tempDict("min"))

Select Case A1 = A2
    
    Case True 'homozygous
        tempRMP = (p * p) + (p * (1 - p) * theta)
        
    Case False 'heterozygous
        tempRMP = 2 * p * q

End Select

CalculateRMP = tempRMP

End Function



'****************************************************
'                cGenotypeSTRmix Properties
'****************************************************
Public Property Get LocusName() As String
    LocusName = pLocusName
End Property
Public Property Let LocusName(v As String)
    pLocusName = v
End Property

Public Property Get Genotype() As String
    Genotype = pGenotype
End Property
Public Property Let Genotype(v As String)
    pGenotype = v
End Property

Public Property Get Percent() As Double
    Percent = pPercent
End Property
Public Property Let Percent(v As Double)
    pPercent = v
End Property

Public Property Get RMP() As Double
    RMP = pRMP
End Property
Public Property Let RMP(v As Double)
    pRMP = v
End Property

Public Property Get Alleles() As Dictionary
    Set Alleles = pAlleles
End Property
Public Property Let Alleles(v As Dictionary)
    Set pAlleles = v
End Property

