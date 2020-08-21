VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCODIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Represents a CODIS sample, containing data to produce a CMF file
'Created from the CODIS sheet data instead of directly from a cContributor object to allow analysts to manually edit

Private pAnalyst As String
Private pDestORI As String
Private pSourceLab As String

Private pCaseID As String
Private pSpecimenID As String
Private pSpecimenCategory As String
Private pSourceID As String
Private pSpecimenComment As String
Private pContributorNum As Integer

'these are arbitrary but not sure we can omit them:
Private pSubmitDateTime As String 'The time this cCODIS object is created (by clicking "Export CMF")
Private pReadingDateTime As String 'This usually = the time you saved the GMID project, but we can't do that here so <shrug>

Private pLocusCategories As New Scripting.Dictionary 'key = locus name, item = partial locus T/F
Private pLocusAlleles As New Scripting.Dictionary 'key = locus name, item = dictionary of alleles with their obligate status

Dim xmlDoc As New MSXML2.DOMDocument60  'the actual xml document

'These values will barely ever change, so they don't need to be a property of every cCODIS object:
Dim HeaderVersion As String
Dim MessageType As String
Dim KitName As String
Dim PartialSpec As Boolean



Public Sub InitializeMe(CODISprofile As Range, SpecifyContributors As Boolean)
'CODISprofile is the "Dest_CODIS1" or "Dest_CODIS2" named range on the CODIS worksheet
'Loads data from CODIS sheet into memory within this cCODIS object
'SpecifyContributors is True when more than one profile from the same sample is being uploaded; the file will get the contributor # tacked onto the name

    Dim thisCODISsheet As Worksheet: Set thisCODISsheet = CODISprofile.parent
    Dim dictAlleles As Dictionary 'key = Allele, value = obligate T/F
    Dim strAlleles As String
    Dim arrAlleles() As String
    Dim rng As Range, v As Variant, i As Integer
    
    Me.Analyst = Sheets("STRlite Settings").Range("CODIS_UserID").Value
    Me.SourceLab = Sheets("STRlite Settings").Range("CODIS_SourceLab").Value
    
    HeaderVersion = Sheets("STRlite Settings").Range("Settings_CODIS").Offset(1, 0).Value
    MessageType = Sheets("STRlite Settings").Range("Settings_CODIS").Offset(2, 0).Value
    Me.DestORI = Sheets("STRlite Settings").Range("Settings_CODIS").Offset(3, 0).Value
    KitName = Sheets("STRlite Settings").Range("Settings_CODIS").Offset(4, 0).Value
    
    Me.ContributorNum = CODISprofile.Cells(1, 1).Offset(-1, 0).Value
    Me.SpecimenComment = thisCODISsheet.Range("CODIS_Comment").Value
    Me.SourceID = GetSourceID(Range(Replace(CODISprofile.Name.Name, "Dest_CODIS", "MatchesKnown")).Value)
    Me.SpecimenCategory = GetSpecCat(CODISprofile.Cells(1, 1).Offset(-2, 1).Value)

    Me.CaseID = thisCODISsheet.Range("CODIS_CaseNum").Value 'this is the case # as the analyst enters it or as it comes out of STRmix
    Me.SubmitDateTime = FormatTimeStamp(Format(Now, "yyyy-mm-dd-hh-mm-ss")) 'THIS date and time. Like, right now.
    Me.ReadingDateTime = FormatTimeStamp(thisCODISsheet.Range("CODIS_DeconTime").Value) 'From the Decon Timestamp
        
    Dim tempSampleName As String
    tempSampleName = PubFun.FixWorksheetName(thisCODISsheet.Range("CODIS_SampleID").Value) 'populated from the decon sampleID, with illegal characters replaced
    If InStr(1, tempSampleName, "Decon", vbTextCompare) > 0 Then 'get rid of the "Decon" that we tack onto the end
        tempSampleName = Trim(Left(tempSampleName, InStr(1, tempSampleName, "Decon", vbTextCompare) - 1)) 'Gets rid of everything past & including the "Decon"
        If Not IsAlphaNumeric(Right(tempSampleName, 1)) Then tempSampleName = Left(tempSampleName, Len(tempSampleName) - 1) 'chop off a trailing non-alphanumeric character
    End If
    
    
    Me.SpecimenID = Me.CaseID & "_" & tempSampleName
    
    If SpecifyContributors Then Me.SpecimenID = Me.SpecimenID & "_" & Me.ContributorNum 'tack on contributor # if necessary

    'Load CODIS profile into Me.LocusCategories and Me.LocusAlleles
    PartialSpec = False
    For Each rng In CODISprofile
        'Check if amelogenin is blank and remind the analyst to fill it in
        If rng.Value = "Amelogenin" And rng.Offset(0, 1).Value = "" Then
            Dim answer As Variant
            answer = MsgBox("Amelogenin is blank. If you can make a call, you must type it in manually first." & vbNewLine & vbNewLine & _
                            "Do you want to proceed without amelogenin?" & vbNewLine & "(Click No to go back and change it)", vbYesNo + vbExclamation, "Determine Amelogenin?")
            If answer = vbNo Then Exit Sub
        End If
        
    
        If rng.Value <> "" And rng.Offset(0, 1).Value <> "" And rng.Offset(0, 1).Value <> "Omitted" Then 'If the locus isn't blank and isn't "Omitted"
        
            'Add locus and its partial (M) status
            Me.LocusCategories.Add rng.Value, IIf(rng.Offset(0, 2).Value = "M", True, False)
                If rng.Offset(0, 2).Value = "M" Then PartialSpec = True 'Any M locus changes PartialSpec flag to True
    
            strAlleles = rng.Offset(0, 1).Value
            strAlleles = Replace(Replace(strAlleles, "{", ""), "}", "") 'get rid of curly brackets if present
            
            arrAlleles = Split(strAlleles, ", ")
            Set dictAlleles = New Scripting.Dictionary
            
            For i = 0 To UBound(arrAlleles)
                If InStr(1, arrAlleles(i), "+", vbTextCompare) > 0 Then
                    dictAlleles.Add Replace(arrAlleles(i), "+", ""), True
                Else: dictAlleles.Add arrAlleles(i), False
                End If
            Next i
                
            Me.LocusAlleles.Add rng.Value, dictAlleles
        End If
    Next rng
    
    
    Call Me.GenerateXML
    
End Sub


Function FormatTimeStamp(InputStamp As String) As String

    Dim temp As String: temp = InputStamp
    Dim tempdate As String, temptime As String
    
    'Fix VarNOC timestamp, which will have a ")_" right before the actual timestamp:
    If InStr(1, temp, ")_", vbTextCompare) > 0 Then temp = Mid(temp, InStr(1, temp, ")_", vbTextCompare) + 2)
        
    tempdate = Left(temp, 10)
    temptime = Replace(Mid(temp, 12), "-", ":")
    
    FormatTimeStamp = tempdate & "T" & temptime
    
End Function

Sub GenerateXML()

    'VBA XML Notes:
    '<ELEMENT attribute="attributevalue">element.text</ELEMENT> Both elements and attributes are types of Nodes
    'xmlDoc.createElement does not allow namespace inheriting, so use xmlDoc.createNode(1,"nodename","namespaceURI") instead.
    'xmlDoc.createNode(1, "nodename", "namespaceURI") 1 means Element node, then the name, then a reference to the namespaceURI.
    
    
    Dim xmlRoot As MSXML2.IXMLDOMElement    'the xml root node. In this document, "CODISImportFile"
    Dim Locus As Variant, allele As Variant

    'Getting started:
        'Write the <?xml version="1.0"?> processing instruction
        xmlDoc.appendChild xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
        
        'Create the root node and add it to xmlDoc
        Set xmlRoot = xmlDoc.createNode(1, "CODISImportFile", "urn:CODISImportFile-schema")
            'urn:CODISImportFile-schema = needs to be the default namespaceURI for the whole xmlDoc, but it doesn't inherit automatically with .createElement method
        xmlDoc.appendChild xmlRoot
        
        AddElementAndValue "HEADERVERSION", HeaderVersion, xmlRoot
        AddElementAndValue "MESSAGETYPE", MessageType, xmlRoot
        AddElementAndValue "DESTINATIONORI", Me.DestORI, xmlRoot
        AddElementAndValue "SOURCELAB", Me.SourceLab, xmlRoot
        AddElementAndValue "SUBMITBYUSERID", Me.Analyst, xmlRoot
        AddElementAndValue "SUBMITDATETIME", Me.SubmitDateTime, xmlRoot
    
        Dim SpecNode As IXMLDOMElement: Set SpecNode = xmlDoc.createNode(1, "SPECIMEN", xmlRoot.NamespaceURI)
        SpecNode.setAttribute "PARTIAL", IIf(PartialSpec, "true", "false")
        SpecNode.setAttribute "CASEID", Left(Me.CaseID, 32) 'Case ID has max 32 characters in CODIS
        SpecNode.setAttribute "SOURCEID", Me.SourceID
        
        AddElementAndValue "SPECIMENID", Left(Me.SpecimenID, 24), SpecNode 'Specimen ID has max 24 characters in CODIS
        AddElementAndValue "SPECIMENCATEGORY", Me.SpecimenCategory, SpecNode
        AddElementAndValue "SPECIMENCOMMENT", Me.SpecimenComment, SpecNode
        
        
    'Locus/Allele looping time!
        Dim LocusNode As IXMLDOMElement
        Dim AlleleNode As IXMLDOMElement
    
        For Each Locus In Me.LocusCategories.Keys
        
            'Create a new LocusNode
            Set LocusNode = xmlDoc.createNode(1, "LOCUS", xmlRoot.NamespaceURI)
            'LocusNode.setAttribute "KIT", KitName
            LocusNode.setAttribute "PARTIALLOCUS", IIf(LocusCategories(Locus), "true", "false")
            
            'Add locus elements:
            AddElementAndValue "LOCUSNAME", CStr(Locus), LocusNode
            AddElementAndValue "READINGBY", CStr(Me.Analyst), LocusNode
            AddElementAndValue "READINGDATETIME", Me.ReadingDateTime, LocusNode
            
            'Create new allele nodes
            Dim tempDict As Scripting.Dictionary: Set tempDict = Me.LocusAlleles(Locus)
            For Each allele In tempDict.Keys
                Set AlleleNode = xmlDoc.createNode(1, "ALLELE", xmlRoot.NamespaceURI)
                If tempDict(allele) = True Then AlleleNode.setAttribute "ALLELEREQUIRED", "true"
                AddElementAndValue "ALLELEVALUE", CStr(allele), AlleleNode
                
                'Add the allele node as a child of the locus node
                LocusNode.appendChild AlleleNode
        
            Next allele
            
            'Add the locus node as a child of the specimen node
            SpecNode.appendChild LocusNode
            
        Next Locus
        
        xmlRoot.appendChild SpecNode

    'Finally, save the xml file to disk
    Dim CMFpath As String


    On Error GoTo PathError
    CMFpath = Sheets("STRlite Settings").Range("CMF_FolderPath").Value
    If CMFpath = "" Then GoTo PathError
SaveCMFPath:
    xmlDoc.Save (CMFpath & "\" & "CMF_" & Me.SpecimenID & ".xml")
    On Error GoTo 0
    
    MsgBox "A CODIS CMF file has been saved to the following location:" & vbNewLine & vbNewLine & _
    CMFpath & "\" & "CMF_" & Me.SpecimenID & ".xml", vbOKOnly + vbInformation, "CMF File Created"
    
    Exit Sub
    
PathError:
    CMFpath = ActiveWorkbook.Path
    GoTo SaveCMFPath
    
End Sub

Sub AddElementAndValue(NodeName As String, NodeValue As String, AddTo As IXMLDOMNode)

    Dim newElement As IXMLDOMElement
    Set newElement = xmlDoc.createNode(1, CStr(NodeName), AddTo.NamespaceURI) 'AddTo.NamespaceURI makes the newElement inherit namespace from its "parent"
    newElement.Text = CStr(NodeValue)
        
    AddTo.appendChild newElement

End Sub


Function GetSpecCat(Category As String) As String

    Select Case Category
    
        Case "Unknown"
            GetSpecCat = "Forensic, Unknown"
            
        Case "Mixture"
            GetSpecCat = "Forensic Mixture"
            
        Case "Partial"
            GetSpecCat = "Forensic Partial"
        
        Case "Targeted"
            GetSpecCat = "Forensic Targeted"
            
        Case Else
            GetSpecCat = Category 'allows for custom category additions to export as-is
        
    End Select

End Function

Function GetSourceID(MatchesKnown As String) As String

    Select Case MatchesKnown
    
        Case "None"
            GetSourceID = "No"
            
        Case ""
            GetSourceID = "No"
        
        Case Else
            GetSourceID = "Yes"
        
    End Select

End Function

Public Property Get Analyst() As String
    Analyst = pAnalyst
End Property
Public Property Let Analyst(v As String)
    pAnalyst = v
End Property

Public Property Get DestORI() As String
    DestORI = pDestORI
End Property
Public Property Let DestORI(v As String)
    pDestORI = v
End Property

Public Property Get SourceLab() As String
    SourceLab = pSourceLab
End Property
Public Property Let SourceLab(v As String)
    pSourceLab = v
End Property



Public Property Get CaseID() As String
    CaseID = pCaseID
End Property
Public Property Let CaseID(v As String)
    pCaseID = v
End Property

Public Property Get SpecimenID() As String
    SpecimenID = pSpecimenID
End Property
Public Property Let SpecimenID(v As String)
    pSpecimenID = v
End Property

Public Property Get SpecimenCategory() As String
    SpecimenCategory = pSpecimenCategory
End Property
Public Property Let SpecimenCategory(v As String)
    pSpecimenCategory = v
End Property

Public Property Get SourceID() As String
    SourceID = pSourceID
End Property
Public Property Let SourceID(v As String)
    pSourceID = v
End Property

Public Property Get SpecimenComment() As String
    SpecimenComment = pSpecimenComment
End Property
Public Property Let SpecimenComment(v As String)
    pSpecimenComment = v
End Property

Public Property Get ContributorNum() As Integer
    ContributorNum = pContributorNum
End Property
Public Property Let ContributorNum(v As Integer)
    pContributorNum = v
End Property



Public Property Get SubmitDateTime() As String
    SubmitDateTime = pSubmitDateTime
End Property
Public Property Let SubmitDateTime(v As String)
    pSubmitDateTime = v
End Property

Public Property Get ReadingDateTime() As String
    ReadingDateTime = pReadingDateTime
End Property
Public Property Let ReadingDateTime(v As String)
    pReadingDateTime = v
End Property


Public Property Get LocusCategories() As Dictionary
    Set LocusCategories = pLocusCategories
End Property
Public Property Let LocusCategories(v As Dictionary)
    Set pLocusCategories = v
End Property

Public Property Get LocusAlleles() As Dictionary
    Set LocusAlleles = pLocusAlleles
End Property
Public Property Let LocusAlleles(v As Dictionary)
    Set pLocusAlleles = v
End Property
