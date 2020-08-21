VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStutterModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pStutterName As String
Private pStutterVariance As String '
Private pStutterMax As Double
Private pStutterEnabled As Boolean
Private pStutterChart As String 'name of stutter variance graph file for this stutter type.
'For v2.5, StutterChart name is fixed. v2.6+ can have custom stutter but the graphs get named Stut1, Stut2, Stut3, etc.



Public Property Get StutterName() As String
    StutterName = pStutterName
End Property
Public Property Let StutterName(v As String)
    pStutterName = v
End Property


Public Property Get StutterVariance() As String
    StutterVariance = pStutterVariance
End Property
Public Property Let StutterVariance(v As String)
    pStutterVariance = v
End Property


Public Property Get StutterMax() As Double
    StutterMax = pStutterMax
End Property
Public Property Let StutterMax(v As Double)
    pStutterMax = v
End Property


Public Property Get StutterChart() As String
    StutterChart = pStutterChart
End Property
Public Property Let StutterChart(v As String)
    pStutterChart = v
End Property


Public Property Get StutterEnabled() As Boolean
    StutterEnabled = pStutterEnabled
End Property
Public Property Let StutterEnabled(v As Boolean)
    pStutterEnabled = v
End Property

