VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLocusWhole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Represents a whole, undeconvoluted (convoluted?) locus
'Source is either the genotype table from GMID or the Evidence File from a STRmix Decon (which itself comes from the GMID genotype table)
'Contains all locus alleles and their peak heights in the Alleles dictionary
'***************************************************************************************************

Private pLocusName As String
Private pAlleles As New Scripting.Dictionary 'key = allele, item = peak height
Private pAlleleString As String
Private pStutterString As String
Private pSaturationFlag As Boolean

Dim allele As Variant
Dim PossibleStutter As New Scripting.Dictionary
Dim NotStutter As New Scripting.Dictionary

Option Explicit


Public Sub InitializeMe(LocusRange As Range, Standard As Boolean)
'Initializes from Genotype table format (GMID)
'LocusRange is the Marker cell of the Genotype Table row for a certain locus
'This is where the stochastic threshold gets applied to standards

    Me.LocusName = LocusRange.Value

    Dim cols As Integer: cols = NumGenotypeTableAlleles 'shortcut
    Dim rng As Range, peakheight As Long
    
    Dim AlleleRange As Range: Set AlleleRange = LocusRange.Offset(0, 1).Resize(1, cols)
    
    For Each rng In AlleleRange
    
        If rng.Value <> "" And rng.Value <> "OL" Then
            peakheight = rng.Offset(0, 2 * cols).Value
            Alleles.Add rng.Value, peakheight
            If SaturationMax <> 0 And peakheight > SaturationMax Then Me.SaturationFlag = True
        End If
        
    Next rng
    
    'Standards get stochastic threshold applied; unknowns get stutter filter applied
    If Standard Then
    
        Select Case Me.Alleles.Count
            
            Case 0
                Me.AlleleString = ""
            
            Case 1
                Me.AlleleString = IIf(CInt(Alleles.Items(0)) < StdStochastic, "[" & Alleles.Keys(0) & "]", Alleles.Keys(0))

            Case 2
                If Alleles.Items(0) < StdStochastic And Alleles.Items(1) < StdStochastic Then
                    Me.AlleleString = "[" & Alleles.Keys(0) & ", " & Alleles.Keys(1) & "]"
                Else:
                    If Alleles.Items(0) < StdStochastic Then Alleles.Keys(0) = "[" & Alleles.Keys(0) & "]"
                    If Alleles.Items(1) < StdStochastic Then Alleles.Keys(1) = "[" & Alleles.Keys(1) & "]"
                    Me.AlleleString = Alleles.Keys(0) & ", " & Alleles.Keys(1)
                End If
            
            Case Is > 2
                Me.AlleleString = "Check"
        
        End Select
        
    Else:
    
        Me.AnalyzeStutter
        
    End If
    
End Sub


Sub AnalyzeStutter()

Dim ratio As Double 'Stutter ratio dictionaries: key = locus, item = ratio
Dim parent As Double
Dim phAllele As Long

If BackStutterRatios Is Nothing Or ForwardStutterRatios Is Nothing Or N2StutterRatios Is Nothing Or DBStutterRatios Is Nothing Then Call Admin.LociCollection
    
'Backward stutter
    If BackStutterRatios.Exists(Me.LocusName) Then
        ratio = BackStutterRatios(Me.LocusName)
        
        For Each allele In Me.Alleles.Keys
            parent = allele + 1
            phAllele = Alleles(allele)
            'If there is an allele one repeat larger, and allele's peak height is < ratio * parent peak's height, add it to PossibleStutter
            If Alleles.Exists(parent) Then
                If phAllele <= CDec(ratio * Alleles(parent)) Then
                    If Not PossibleStutter.Exists(allele) Then PossibleStutter.Add allele, allele
                End If
            End If
        Next allele
    End If
                                
'Forward stutter
    If ForwardStutterRatios.Exists(Me.LocusName) Then
        ratio = ForwardStutterRatios(Me.LocusName)
        
        For Each allele In Me.Alleles.Keys
            parent = allele - 1
            phAllele = Alleles(allele)
            'If there is an allele one repeat smaller, and allele's peak height is < ratio * parent peak's height, add it to PossibleStutter
            If Alleles.Exists(parent) Then
                If phAllele <= CDec(ratio * Alleles(parent)) Then
                    If Not PossibleStutter.Exists(allele) Then PossibleStutter.Add allele, allele
                End If
            End If
        Next allele
    End If
    
    
'N-2 stutter. Bit more complicated, but should only be needed for a few loci. Only works for 4 bp repeats.
    If N2StutterRatios.Exists(Me.LocusName) Then
        ratio = N2StutterRatios(Me.LocusName)
        Dim FractionalAlleles As New Scripting.Dictionary
        Dim fractAllele As Variant
        
        'Convert alleles to decimals & store in FractionalAlleles
        For Each allele In Me.Alleles.Keys
            fractAllele = PubFun.AlleleToDecimal(CDbl(allele), 4) 'Using 4 as the repeat size
            If Not FractionalAlleles.Exists(allele) Then FractionalAlleles.Add fractAllele, Me.Alleles(allele) 'Me.Alleles(allele)= peak height
        Next allele
            
        'Now we can define n-2 stutter as 1/2 of one repeat, i.e. 0.5
        For Each allele In FractionalAlleles.Keys
            If FractionalAlleles.Exists(CDbl(allele) + 0.5) Then
                If FractionalAlleles(allele) <= CDec(ratio * FractionalAlleles(allele + 0.5)) Then
                    fractAllele = DecimalToAllele(CDbl(allele), 4)
                    If Not PossibleStutter.Exists(fractAllele) Then PossibleStutter.Add CVar(fractAllele), fractAllele
                End If
            End If
        Next allele
    End If
    
'Double-back stutter
    If DBStutterRatios.Exists(Me.LocusName) Then
        ratio = DBStutterRatios(Me.LocusName)
        
        For Each allele In Me.Alleles.Keys
            parent = allele + 2
            phAllele = Alleles(allele)
            'If there is a parent two repeats larger, and allele's peak height is < ratio * parent peak's height, add it to PossibleStutter
            If Alleles.Exists(parent) Then
                If phAllele <= CDec(ratio * Alleles(parent)) Then
                    If Not PossibleStutter.Exists(allele) Then PossibleStutter.Add allele, allele
                End If
            End If
        Next allele
    End If
    
    
'Sort alleles and turn into strings

Dim tempAlleles() As Variant
Dim tempStutter() As Variant
    
    For Each allele In Me.Alleles.Keys
        If Not PossibleStutter.Exists(allele) Then NotStutter.Add allele, Nothing
    Next allele

    tempAlleles = NotStutter.Keys()
    Me.AlleleString = Join(tempAlleles, ", ")
    
    tempStutter = PossibleStutter.Keys()
    Me.StutterString = Join(tempStutter, ", ")

End Sub


Sub InitializeFromSTRmixEV(LocusName As String)
'INCOMPLETE/NOT CURRENTLY USED
'STRmix isolates the allele data in a different format from the GMID-style Genotype Table.
'This harvests the locus from a STRmix Evidence Input File which has already been copied into Import tab

Me.LocusName = LocusName

Dim LocusStart As Range: LocusStart = FindCell(Sheets("Import").Range("A1:A600"), LocusName, Sheets("Import").Range("A1"))
Dim AlleleRow As Range: Set AlleleRow = LocusStart

Do While AlleleRow.Value = LocusName
    With AlleleRow
        Alleles.Add .Offset(0, 1).Value, .Offset(0, 2).Value
    End With
    
    Set AlleleRow = AlleleRow.Offset(1, 0)
Loop

End Sub



Public Property Get LocusName() As String
    LocusName = pLocusName
End Property
Public Property Let LocusName(v As String)
    pLocusName = v
End Property

Public Property Get Alleles() As Dictionary
    Set Alleles = pAlleles
End Property
Public Property Let Alleles(v As Dictionary)
    Set pAlleles = v
End Property

Public Property Get AlleleString() As String
    AlleleString = pAlleleString
End Property
Public Property Let AlleleString(v As String)
    pAlleleString = v
End Property

Public Property Get StutterString() As String
    StutterString = pStutterString
End Property
Public Property Let StutterString(v As String)
    pStutterString = v
End Property

Public Property Get SaturationFlag() As Boolean
    SaturationFlag = pSaturationFlag
End Property
Public Property Let SaturationFlag(v As Boolean)
    pSaturationFlag = v
End Property

