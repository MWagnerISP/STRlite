VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProfileGenotype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pSampleName As String
Private pLoci As New Scripting.Dictionary 'key = locus name, item = cLocusWhole
Private pIsStandard As Boolean
Private pSaturation As Boolean 'flag for detection of oversaturated peak(s)



Sub DumpData(TargetSheet As Worksheet, Optional Conditioned As Boolean = False, Optional col As Integer)
'Cond is used for a conditioned standard. col is supplied if the sample is a regular standard

Dim Dest As Range, rng As Range

    Select Case Me.IsStandard
    
        Case True
            If Conditioned Then
                'Conditioned standard goes on PreSTRmix sheet (TargetSheet)
                TargetSheet.Range("F7").Value = Me.SampleName
                Set Dest = TargetSheet.Range("Dest_LociPreSTRmix")
                col = 4
            Else:
                'Regular standard goes on Standard sheet (TargetSheet)
                TargetSheet.Range("Dest_StandardSampleName").Offset(0, col).Value = Me.SampleName
                Set Dest = TargetSheet.Range("Dest_LociStandards")
            End If
            
            For Each rng In Dest
                If Me.Loci.Exists(rng.Value) Then
                    rng.Offset(0, col).Value = Me.Loci(rng.Value).AlleleString
                End If
            Next rng
            
        Case False
            TargetSheet.Range("K4").Value = Me.SampleName
            Set Dest = TargetSheet.Range("Dest_LociPreSTRmix")
            
            For Each rng In Dest
                If Me.Loci.Exists(rng.Value) Then
                    rng.Offset(0, 1).Value = Me.Loci(rng.Value).AlleleString
                    rng.Offset(0, 2).Value = Me.Loci(rng.Value).StutterString
                End If
            Next rng
        
    End Select

End Sub


Public Property Get SampleName() As String
    SampleName = pSampleName
End Property
Public Property Let SampleName(v As String)
    pSampleName = v
End Property

Public Property Get Loci() As Dictionary
    Set Loci = pLoci
End Property
Public Property Let Loci(v As Dictionary)
    Set pLoci = v
End Property

Public Property Get IsStandard() As Boolean
    IsStandard = pIsStandard
End Property
Public Property Let IsStandard(v As Boolean)
    pIsStandard = v
End Property

Public Property Get Saturation() As Boolean
    Saturation = pSaturation
End Property
Public Property Let Saturation(v As Boolean)
    pSaturation = v
End Property
