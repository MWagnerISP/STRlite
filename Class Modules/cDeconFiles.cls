VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDeconFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Files & Folders
Private pSTRmixVersion As String

Private pFileName_Decon As String 'Full name of Results.txt file, should be CaseID_Results.txt
Private pFilePath_Decon As String 'Filepath to Results.txt file, navigated by user
Private pFolderPath_Decon As String 'Full path of Decon folder containing Results.txt file
Private pFolderName_Decon As String 'Full name of Decon folder
Private pDeconFolderBasename As String 'Decon folder name without timestamp
Private pFolderPath_Parent As String 'Main folder filepath (one up from Decon & DBSearch folders)
Private pFolderName_Parent As String 'Name of folder containing decon & DBsearch folders
Private pFolderPath_Reports As String '"Reports" folder, contains CODIS xmls, PDF, & variance graphs
Private pFilePath_Contributor As String 'Filepath of ComponentInterp file
Private pFileName_Contributor As String 'File name of ComponentInterp file
Private pFilePath_DeconConfigXML As String 'Filepath of Config.xml file
Private pFilePath_DeconResultsXML As String 'Filepath of Results.xml file
Private pFilePath_SearchResultXML As String 'Filepath of Database search results.xml file
Private pFilePath_AlleleChart As String
Private pFilePath_degCurveChart As String

Private pFilePath_StutterCharts As New Scripting.Dictionary 'key = chart file name; item = filepath of the variance chart

'Private pFilePath_BackStutterChart As String
'Private pFilePath_ForwardStutterChart As String
'Private pFilePath_SE33StutterChart As String


Option Explicit



Public Sub InitializeMe(DeconFolder As String, NOC As Integer, Version As String)
'NOC must be specified when it's a VarNOC decon

    Dim fsoDecon As FileSystemObject
    Set fsoDecon = CreateObject("Scripting.FileSystemObject")
    
    Dim i As Integer
    
    'Shortcut variables
    Dim ReportFolder As String
    Dim DeconFolderName As String
    
    
    With Me
    
    'STRmix Version
        .STRmixVersion = Version
    
    'DeconFolder holds everything
        .FolderPath_Decon = DeconFolder 'Full path of Decon folder containing Results.txt file
        .FolderName_Decon = fsoDecon.GetFileName(DeconFolder) 'Full name of Decon folder
        DeconFolderName = .FolderName_Decon
    
    'Parent folder
        .FolderPath_Parent = fsoDecon.GetParentFolderName(DeconFolder) 'Parent folder filepath (one up from Decon & DBSearch folders)
        .FolderName_Parent = fsoDecon.GetFileName(.FolderPath_Parent) 'Name of parent folder containing decon & DBsearch folders
        
    'Results.txt file
        .FileName_Decon = Dir(DeconFolder & "/*Results.txt") 'Full name of Results.txt file, should be CaseID_Results.txt
        .FilePath_Decon = DeconFolder & "/" & .FileName_Decon 'path of Results.txt file
        
    'XML files
        .FilePath_DeconConfigXML = DeconFolder & "/config.xml" 'Filepath of Config.xml file
        .FilePath_DeconResultsXML = DeconFolder & "/results.xml" 'Filepath of Results.xml file
        
    'Split the Decon Folder name starting with the timestamp, which starts with the year. Should work through 2039...
        On Error Resume Next
            If InStr(1, DeconFolderName, "_201") > 0 Then .DeconFolderBasename = Left(DeconFolderName, InStrRev(DeconFolderName, "_201") - 1)
            If InStr(1, DeconFolderName, "_202") > 0 Then .DeconFolderBasename = Left(DeconFolderName, InStrRev(DeconFolderName, "_202") - 1)
            If InStr(1, DeconFolderName, "_203") > 0 Then .DeconFolderBasename = Left(DeconFolderName, InStrRev(DeconFolderName, "_203") - 1)
        On Error GoTo 0
        
    'FilePath_SearchResult  'Filepath of Database search results.xml file
        If Sheets("STRlite Settings").Range("DBSearchOn").Value = True Then 'can be turned off on the STRlite Settings tab
            If Dir(.FolderPath_Parent & "\" & .DeconFolderBasename & "-DBSearch*", vbDirectory) = "" Then 'If the DBSearch folder isn't where it should be...
                Dim answer As String
                Dim FindMe As FileDialog: Set FindMe = Application.FileDialog(msoFileDialogFolderPicker)
                answer = MsgBox("STRlite can't find the DBSearch folder for this Deconvolution." & vbNewLine & "Please select the DBSearch folder.", _
                vbOKCancel + vbCritical, "DBSearch Folder Not Found")
                If answer = vbOK Then
                    FindMe.AllowMultiSelect = False 'doesn't work for FolderPicker anyway
                    FindMe.Title = "Please select the DBSearch folder for this Deconvolution."
                    FindMe.ButtonName = "Select Folder"
                    FindMe.InitialFileName = .FolderPath_Parent
                    If FindMe.Show <> 0 Then
                        .FilePath_SearchResultXML = FindMe.SelectedItems(1) & "\results.xml"
                    Else: .FilePath_SearchResultXML = "n/a"
                    End If
                End If
                If answer = vbCancel Then .FilePath_SearchResultXML = "n/a"
            Else:
                .FilePath_SearchResultXML = .FolderPath_Parent & "\" & Dir(.FolderPath_Parent & "\" & .DeconFolderBasename & "-DBSearch*", vbDirectory) & "\results.xml"
            End If
        End If
        
    'Reports folder
    If fsoDecon.FolderExists(DeconFolder & "/Reports") Then
        .FolderPath_Reports = DeconFolder & "/Reports"
        ReportFolder = .FolderPath_Reports & "/"
        Else: MsgBox "Error: Reports folder not found.", vbOKOnly, "Missing Reports Folder"
    End If
        
    'These files vary/multiply with VarNOC:
        
        If NOC = 0 Then '(No VarNOC)
    
            'Component Interpretation file
                .FileName_Contributor = Dir(DeconFolder & "/" & "*_ComponentInterp*.txt")
                .FilePath_Contributor = DeconFolder & "/" & .FileName_Contributor
            
            'Graphs within ReportFolder
                
                If .STRmixVersion = "2.5" Then 'v2.5 has different names for these files (and fewer of them)
                    .FilePath_AlleleChart = ReportFolder & Dir(ReportFolder & "*alleleVarianceChart.png") 'Allele Variance Chart
                    FilePath_StutterCharts.Add "Back Stutter", ReportFolder & Dir(ReportFolder & "*stutterVarianceChart.png") 'Stutter Variance Chart
                    
                Else: 'For v2.6+
                    .FilePath_AlleleChart = ReportFolder & Dir(ReportFolder & "*Var_allele.png")
                    .FilePath_degCurveChart = IIf(Dir(ReportFolder & "*degCurve.png") <> "", ReportFolder & Dir(ReportFolder & "*degCurve.png"), "None")
                    
                    For i = 1 To 4 'Up to 4 types of stutter
                        If Dir(ReportFolder & "*Var_Stut" & i & ".png") <> "" Then FilePath_StutterCharts.Add "Var_Stut" & i, ReportFolder & Dir(ReportFolder & "*Var_Stut" & i & ".png")
                        'Debug.Print FilePath_StutterCharts("Var_Stut" & i)
                    Next i
                    
                End If
                
            
            
        Else: 'for VarNOC, NOC is specified and thus <> 0
        
            'Component Interpretation file
                .FileName_Contributor = Dir(DeconFolder & "/" & "*_ComponentInterp" & NOC & "T.txt") 'grab the one for this NOC
                .FilePath_Contributor = DeconFolder & "/" & .FileName_Contributor
            
            'Graphs within ReportFolder: different naming convention from non-VarNOC files
                .FilePath_AlleleChart = ReportFolder & Dir(ReportFolder & "*" & NOC & "cont) Var_allele.png")
                .FilePath_degCurveChart = ReportFolder & Dir(ReportFolder & "*" & NOC & "C)_degCurve.png") 'degCurve png file has a different naming convention
                
                For i = 1 To 4 'Up to 4 types of stutter
                    If Dir(ReportFolder & "*" & NOC & "cont) Var_Stut" & i & ".png") <> "" Then FilePath_StutterCharts.Add "Var_Stut" & i, ReportFolder & Dir(ReportFolder & "*" & NOC & "cont) Var_Stut" & i & ".png")
                Next i
                
        End If
            

    Set fsoDecon = Nothing
    
    End With

End Sub



'Files and Folders

Public Property Get STRmixVersion() As String
    STRmixVersion = pSTRmixVersion
End Property
Public Property Let STRmixVersion(v As String)
    pSTRmixVersion = v
End Property

Public Property Get FileName_Decon() As String
    FileName_Decon = pFileName_Decon
End Property
Public Property Let FileName_Decon(v As String)
    pFileName_Decon = v
End Property

Public Property Get FilePath_Decon() As String
    FilePath_Decon = pFilePath_Decon
End Property
Public Property Let FilePath_Decon(v As String)
    pFilePath_Decon = v
End Property

Public Property Get FolderName_Decon() As String
    FolderName_Decon = pFolderName_Decon
End Property
Public Property Let FolderName_Decon(v As String)
    pFolderName_Decon = v
End Property

Public Property Get FolderPath_Decon() As String
    FolderPath_Decon = pFolderPath_Decon
End Property
Public Property Let FolderPath_Decon(v As String)
    pFolderPath_Decon = v
End Property

Public Property Get FolderName_Parent() As String
    FolderName_Parent = pFolderName_Parent
End Property
Public Property Let FolderName_Parent(v As String)
    pFolderName_Parent = v
End Property

Public Property Get FolderPath_Parent() As String
    FolderPath_Parent = pFolderPath_Parent
End Property
Public Property Let FolderPath_Parent(v As String)
    pFolderPath_Parent = v
End Property

Public Property Get DeconFolderBasename() As String
    DeconFolderBasename = pDeconFolderBasename
End Property
Public Property Let DeconFolderBasename(v As String)
    pDeconFolderBasename = v
End Property

Public Property Get FolderPath_Reports() As String
    FolderPath_Reports = pFolderPath_Reports
End Property
Public Property Let FolderPath_Reports(v As String)
    pFolderPath_Reports = v
End Property


Public Property Get FilePath_Contributor() As String
    FilePath_Contributor = pFilePath_Contributor
End Property
Public Property Let FilePath_Contributor(v As String)
    pFilePath_Contributor = v
End Property

Public Property Get FileName_Contributor() As String
    FileName_Contributor = pFileName_Contributor
End Property
Public Property Let FileName_Contributor(v As String)
    pFileName_Contributor = v
End Property

Public Property Get FilePath_DeconConfigXML() As String
    FilePath_DeconConfigXML = pFilePath_DeconConfigXML
End Property
Public Property Let FilePath_DeconConfigXML(v As String)
    pFilePath_DeconConfigXML = v
End Property

Public Property Get FilePath_DeconResultsXML() As String
    FilePath_DeconResultsXML = pFilePath_DeconResultsXML
End Property
Public Property Let FilePath_DeconResultsXML(v As String)
    pFilePath_DeconResultsXML = v
End Property

Public Property Get FilePath_SearchResultXML() As String
    FilePath_SearchResultXML = pFilePath_SearchResultXML
End Property
Public Property Let FilePath_SearchResultXML(v As String)
    pFilePath_SearchResultXML = v
End Property



Public Property Get FilePath_AlleleChart() As String
    FilePath_AlleleChart = pFilePath_AlleleChart
End Property
Public Property Let FilePath_AlleleChart(v As String)
    pFilePath_AlleleChart = v
End Property


Public Property Get FilePath_StutterCharts() As Dictionary
    Set FilePath_StutterCharts = pFilePath_StutterCharts
End Property
Public Property Let FilePath_StutterCharts(v As Dictionary)
    Set pFilePath_StutterCharts = v
End Property



'Public Property Get FilePath_BackStutterChart() As String
'    FilePath_BackStutterChart = pFilePath_BackStutterChart
'End Property
'Public Property Let FilePath_BackStutterChart(v As String)
'    pFilePath_BackStutterChart = v
'End Property
'
'Public Property Get FilePath_ForwardStutterChart() As String
'    FilePath_ForwardStutterChart = pFilePath_ForwardStutterChart
'End Property
'Public Property Let FilePath_ForwardStutterChart(v As String)
'    pFilePath_ForwardStutterChart = v
'End Property
'
'Public Property Get FilePath_SE33StutterChart() As String
'    FilePath_SE33StutterChart = pFilePath_SE33StutterChart
'End Property
'Public Property Let FilePath_SE33StutterChart(v As String)
'    pFilePath_SE33StutterChart = v
'End Property

Public Property Get FilePath_degCurveChart() As String
    FilePath_degCurveChart = pFilePath_degCurveChart
End Property
Public Property Let FilePath_degCurveChart(v As String)
    pFilePath_degCurveChart = v
End Property



