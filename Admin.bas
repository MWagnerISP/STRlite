Attribute VB_Name = "Admin"
Public LociOrder As Scripting.Dictionary  'contains ALL kit loci in order: key=locus, item = numerical order
Public LociSTRmix As Scripting.Dictionary  'contains all loci typically analyzed by STRmix (i.e. omitting Y loci). Key = locus, value = dictionary of allele frequencies

Public LociSDIS As Scripting.Dictionary 'contains only loci marked for SDIS MRE calculation

Public CODISCore13 As Scripting.Dictionary  'contains CODIS core 13 loci
Public CODISCore20 As Scripting.Dictionary  'contains CODIS core 20 loci

Public CODISMaxAlleles As Integer 'maximum number of alleles per pulled locus for CODIS
Public CODISThreshold As Double 'weight threshold for calling a CODIS locus
Public NumGenotypeTableAlleles As Integer 'number of alleles exported in the Genotype Table
Public SaturationMax As Long 'RFU for saturation flag
Public DegradationFlag As Double 'point at which a contributor gets flagged for high degradation curve D-value
Public StdStochastic As Long 'Stochastic threshold for loci in standards
Public CODISgoalMME As Double 'minimum MME for NDIS
Public NumberWords As Scripting.Dictionary 'key = exponent (integer), item = word (e.g. "trillion")

Public BackStutterRatios As Scripting.Dictionary  'key = locus, item = stutter %
Public ForwardStutterRatios As Scripting.Dictionary  'key = locus, item = stutter %
Public N2StutterRatios As Scripting.Dictionary  'key = locus, item = stutter %
Public DBStutterRatios As Scripting.Dictionary  'key = locus, item = stutter %

'Lists of the (non-template) worksheets that exist at any given time
Public AllSheets As Scripting.Dictionary  'key = sheet name, item = sheet type
Public GMIDSheets As Scripting.Dictionary
Public DeconSheets As Scripting.Dictionary
Public LRSheets As Scripting.Dictionary
Public SingleSheets As Scripting.Dictionary
Public DoubleSheets As Scripting.Dictionary
Public StandardSheets As Scripting.Dictionary
Public CODISSheets As Scripting.Dictionary

Dim ws As Worksheet, i As Integer

Public STRliteUnlocked As Boolean
Public Const theta As Double = 0.01
Public Const STRlitePW As String = "helix" 'shhhh it's a secret.
'To change the password, unlock everything with the current password (including VBA itself), then change STRlitePW, then run LockAll

Option Explicit


Sub AdminForm()
Attribute AdminForm.VB_ProcData.VB_Invoke_Func = "A\n14"
' Shortcut: Ctrl+Shift+A

    FormAdmin.Show

End Sub

Sub GotoMaster()
Attribute GotoMaster.VB_ProcData.VB_Invoke_Func = "m\n14"
' Shortcut: Ctrl+M
'The only reason this is here all by itself is so it can have a macro shortcut (ctrl M)
    Sheets("Master").Select

End Sub

Sub LociCollection()

    Dim rng As Range, tempDictionary As Dictionary
    Application.ScreenUpdating = False
    Sheets("STRlite Settings").Visible = True
    Sheets("NIST 2017").Visible = True
        
    'To avoid multiple sets of loci, start over:
    Set LociSTRmix = Nothing
    Set LociOrder = Nothing
    Set LociSDIS = Nothing
    Set CODISCore13 = Nothing
    Set CODISCore20 = Nothing
    Set BackStutterRatios = Nothing
    Set ForwardStutterRatios = Nothing
    Set N2StutterRatios = Nothing
    Set DBStutterRatios = Nothing
    
    'Fill LociOrder Dictionary (with numerical order)
    i = 1
    Set LociOrder = New Scripting.Dictionary
    Set BackStutterRatios = New Scripting.Dictionary
    Set ForwardStutterRatios = New Scripting.Dictionary
    Set N2StutterRatios = New Scripting.Dictionary
    Set DBStutterRatios = New Scripting.Dictionary
    Set LociSDIS = New Scripting.Dictionary
    
        For Each rng In Sheets("STRlite Settings").Range("Loci_Kit")
            LociOrder.Add rng.Value, i
            
            'Fill the stutter dictionaries while we're at it, because they're right there
            If rng.Offset(0, 1) <> "" Then BackStutterRatios.Add rng.Value, rng.Offset(0, 1).Value
            If rng.Offset(0, 2) <> "" Then ForwardStutterRatios.Add rng.Value, rng.Offset(0, 2).Value
            If rng.Offset(0, 3) <> "" Then N2StutterRatios.Add rng.Value, rng.Offset(0, 3).Value
            If rng.Offset(0, 4) <> "" Then DBStutterRatios.Add rng.Value, rng.Offset(0, 4).Value
            
            'Also detect the SDIS eligible loci by their bold font:
                '(and don't include Y loci or Amelogenin, just in case)
            If rng.Font.Bold And InStr(1, rng.Value, "Y") = 0 And InStr(1, rng.Value, "Amel", vbTextCompare) = 0 Then LociSDIS.Add rng.Value, rng.Value
            
            i = i + 1
        Next rng
        
    'Fill LociSTRmix dictionary
    Set LociSTRmix = New Scripting.Dictionary
        For Each rng In Sheets("STRlite Settings").Range("Loci_STRmix")
            'Store allele frequencies with the STRmix loci
            Set tempDictionary = LoadAlleleFrequencies(rng.Value)
            LociSTRmix.Add rng.Value, tempDictionary
        Next rng
        
    'Fill CODIS Core13 dictionary
    Set CODISCore13 = New Scripting.Dictionary
        For Each rng In Sheets("STRlite Settings").Range("Loci_CODIS13")
            CODISCore13.Add rng.Value, rng.Value
        Next rng
        
    'Fill CODIS Core20 dictionary
    Set CODISCore20 = New Scripting.Dictionary
        For Each rng In Sheets("STRlite Settings").Range("Loci_CODIS20")
            CODISCore20.Add rng.Value, rng.Value
        Next rng
            
            
    Sheets("STRlite Settings").Visible = xlVeryHidden
    Sheets("NIST 2017").Visible = xlVeryHidden
    Application.ScreenUpdating = True
End Sub

Sub LoadSTRliteSettings()

    Application.ScreenUpdating = False
    Sheets("STRlite Settings").Visible = True

    NumGenotypeTableAlleles = Sheets("STRlite Settings").Range("NumGenTableAlleles").Value
    SaturationMax = Sheets("STRlite Settings").Range("SaturationMax").Value
    DegradationFlag = Sheets("STRlite Settings").Range("DegradationFlag").Value
    StdStochastic = Sheets("STRlite Settings").Range("Stochastic").Value
    CODISThreshold = Sheets("STRlite Settings").Range("CODIS_Threshold").Value
    CODISMaxAlleles = Sheets("STRlite Settings").Range("CODIS_MaxAlleles").Value
    CODISgoalMME = Sheets("STRlite Settings").Range("CODIS_goalMME").Value
    
    Set NumberWords = New Scripting.Dictionary
    Dim rng As Range
    For Each rng In Sheets("STRlite Settings").Range("textconversion")
        NumberWords.Add rng.Value, rng.Offset(0, 1).Value
    Next rng
    
    Sheets("STRlite Settings").Visible = xlVeryHidden
    Application.ScreenUpdating = True
    
End Sub

Sub LoadSheetLists()

    Set AllSheets = New Scripting.Dictionary
    Set GMIDSheets = New Scripting.Dictionary
    Set DeconSheets = New Scripting.Dictionary
    Set LRSheets = New Scripting.Dictionary
    Set SingleSheets = New Scripting.Dictionary
    Set DoubleSheets = New Scripting.Dictionary
    Set StandardSheets = New Scripting.Dictionary
    Set CODISSheets = New Scripting.Dictionary
    
    For Each ws In Application.Worksheets
    
        Select Case Left(ws.Name, 4)
            Case "(P) "
                If Not GMIDSheets.Exists(ws.Name) Then GMIDSheets.Add ws.Name, "PreSTRmix"
                If Not AllSheets.Exists(ws.Name) Then AllSheets.Add ws.Name, "PreSTRmix"
            Case "(D) "
                If Not DeconSheets.Exists(ws.Name) Then DeconSheets.Add ws.Name, "Decon"
                If Not AllSheets.Exists(ws.Name) Then AllSheets.Add ws.Name, "Decon"
            Case "(LR)"
                If Not LRSheets.Exists(ws.Name) Then LRSheets.Add ws.Name, "LR"
                If Not AllSheets.Exists(ws.Name) Then AllSheets.Add ws.Name, "LR"
            Case "(Std"
                If Not StandardSheets.Exists(ws.Name) Then StandardSheets.Add ws.Name, "Standard"
                If Not AllSheets.Exists(ws.Name) Then AllSheets.Add ws.Name, "Standard"
            Case "(1P)"
                If Not SingleSheets.Exists(ws.Name) Then SingleSheets.Add ws.Name, "1P"
                If Not AllSheets.Exists(ws.Name) Then AllSheets.Add ws.Name, "1P"
            Case "(2P)"
                If Not DoubleSheets.Exists(ws.Name) Then DoubleSheets.Add ws.Name, "2P"
                If Not AllSheets.Exists(ws.Name) Then AllSheets.Add ws.Name, "2P"
            Case "(C) "
                If Not CODISSheets.Exists(ws.Name) Then CODISSheets.Add ws.Name, "CODIS"
                If Not AllSheets.Exists(ws.Name) Then AllSheets.Add ws.Name, "CODIS"
            
        End Select
        
    Next ws
    
    'Sort the sets of worksheets
    Set GMIDSheets = SortDictionaryByKey(GMIDSheets)
    Set DeconSheets = SortDictionaryByKey(DeconSheets)
    Set LRSheets = SortDictionaryByKey(LRSheets)
    Set SingleSheets = SortDictionaryByKey(SingleSheets)
    Set DoubleSheets = SortDictionaryByKey(DoubleSheets)
    Set StandardSheets = SortDictionaryByKey(StandardSheets)
    Set CODISSheets = SortDictionaryByKey(CODISSheets)

End Sub


Function SortAllSheets(SortType As String) As Object
'returns a dictionary (based on AllSheets) sorted by either worksheet type or case name

    Dim arrName() As String, Name_Type As String, Type_Name As String
    
    Dim sortbyNames As New Scripting.Dictionary 'key= Name_Type, item = original name
    Dim sortbyTypes As New Scripting.Dictionary 'key = Type_Name, item = original name
    Dim arrList As Object: Set arrList = CreateObject("System.Collections.ArrayList")
    
    Dim sortedDict As New Scripting.Dictionary
    
    Dim v As Variant
    
    For Each v In AllSheets.Keys
        arrName = Split(v, " ", 2) 'all STRlite sheet names start with a "(x) " so we split the name after that space
        Name_Type = arrName(1) & "_" & PrefixOrder(arrName(0)) 'enumerates the prefix and puts it at the end. Still unique.
        Type_Name = PrefixOrder(arrName(0)) & "_" & arrName(1) 'enumerates the prefix and keeps it at the beginning. Still unique.
        sortbyNames.Add Name_Type, v
        sortbyTypes.Add Type_Name, v
    Next v
    
    Select Case SortType
        Case "Type"
        
            For Each v In sortbyTypes.Keys
                arrList.Add v
            Next v
            
            arrList.Sort
        
            For Each v In arrList
                sortedDict.Add sortbyTypes(v), Left(v, 1)
            Next v
        
        
        Case "Case"
    
            For Each v In sortbyNames.Keys
                arrList.Add v
            Next v
            
            arrList.Sort
            
            For Each v In arrList
                sortedDict.Add sortbyNames(v), Right(v, 1)
            Next v
            
    End Select
    
    Set arrList = Nothing
    Set SortAllSheets = sortedDict
    
End Function

Function PrefixOrder(Prefix As String) As Integer

    Select Case Prefix
        Case "(P)"
            PrefixOrder = 1
        Case "(D)"
            PrefixOrder = 2
        Case "(1P)"
            PrefixOrder = 3
        Case "(2P)"
            PrefixOrder = 4
        Case "(Std)"
            PrefixOrder = 5
        Case "(LR)"
            PrefixOrder = 6
        Case "(C)"
            PrefixOrder = 7
    End Select

End Function

Function LoadAlleleFrequencies(Locus As String) As Dictionary

    Dim rng As Range, r As Integer, col As Long
    Dim rngAlleles As Range: Set rngAlleles = Sheets("NIST 2017").Range("NIST_Alleles")
    
    col = FindCell(Sheets("NIST 2017").Range("NIST_Loci"), Locus, Range("A5")).Column
    
    'rngLocus is the locus column containing the actual frequencies
    Dim rngLocus As Range: Set rngLocus = rngAlleles.Offset(0, col - 1)
    
    Set LoadAlleleFrequencies = New Scripting.Dictionary
    
    'Iterate through rngAlleles and check if its corresponding frequency in rngLocus exists
    For r = 1 To rngAlleles.Cells.Count
        'Round frequencies to 4 decimal places because that's what CODIS output has
        If rngLocus.Cells(r, 1).Value <> "" Then LoadAlleleFrequencies.Add rngAlleles.Cells(r, 1).Value, Round(CDbl(rngLocus.Cells(r, 1).Value), 4)
    Next
    
End Function


Sub CrackTheHood(Optional Goal As String)

    Application.ScreenUpdating = False
    
        Sheets("Import").Visible = True
        Sheets("STRlite Settings").Visible = True
    
    Select Case Goal
        
        Case "Decon"
            Sheets("Decon Template").Visible = True
    
        Case "GMID"
            Sheets("Pre-STRmix Template").Visible = True
            Sheets("Standards Template").Visible = True
    
        Case "LR"
            Sheets("LR Template").Visible = True
            
        Case "Combo"
            Sheets("1P Template").Visible = True
            Sheets("2P Template").Visible = True
            
        Case "CODIS"
            Sheets("CODIS Template").Visible = True
            Sheets("NIST 2017").Visible = True
            
    End Select

End Sub

Sub CleanUp(Optional TargetSheet As Worksheet)

    Application.ScreenUpdating = False

    Sheets("STRlite Settings").Visible = xlHidden
    Sheets("Import").Visible = xlVeryHidden
    
    Sheets("Pre-STRmix Template").Visible = xlVeryHidden
    Sheets("Standards Template").Visible = xlVeryHidden
    
    Sheets("Decon Template").Visible = xlVeryHidden
            
    Sheets("LR Template").Visible = xlVeryHidden
            
    Sheets("1P Template").Visible = xlVeryHidden
    Sheets("2P Template").Visible = xlVeryHidden
    
    Sheets("CODIS Template").Visible = xlVeryHidden
    Sheets("NIST 2017").Visible = xlVeryHidden
    
    
    If TargetSheet Is Nothing Then
        Sheets("Master").Select
    Else: TargetSheet.Select
    End If
    
    Application.ScreenUpdating = True

End Sub


Sub LockAll()

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "Import" And ws.Name <> "NIST 2017" Then
            With ws
            .EnableSelection = xlUnlockedCells
            .Protect password:=STRlitePW, UserInterfaceOnly:=True, AllowSorting:=True, AllowFormattingCells:=True
            End With
        End If
    Next ws
    
    STRliteUnlocked = False
    
End Sub


Sub FixStupidButtons()
'Fix incredible growing ActiveX Buttons :-(
'This is an Excel/ActiveX bug that is related to changing between monitors/displays. Loves to happen during public presentations...
'See https://stackoverflow.com/questions/1573349/excel-the-incredible-shrinking-and-expanding-controls
'This sub re-defines all the buttons' sizes and then forces them to refresh (see Sub ReScale)

    Application.ScreenUpdating = False
    
    Dim shp As Shape, obj As OLEObject, wks As Worksheet
    'All of these controls are ActiveX objects, which can be both Shapes and OLEObjects (and lots of other things) depending on how you define them.
    'I'm using the Shape type so that I can use the Shape.ScaleHeight/.ScaleWidth methods,
    'and also the OLEObject type to reset the button font size.
    
    'I also discovered that this sub only works for ungrouped controls
    '(otherwise it will see the group and not the individual controls)
    

    'Master tab
    For Each shp In Sheets("Master").Shapes
        Select Case Left(shp.Name, 3)
        
            Case "cmd"
                'Set height/width of Shape
                shp.Height = Application.InchesToPoints(0.5)
                shp.Width = Application.InchesToPoints(1)
                'Set font of OLEObject
                Sheets("Master").OLEObjects(shp.Name).Object.Font.Size = 12
                ReScale shp 'I do this last because it acts like a "refresh"
            
            Case "LB_"
                shp.Height = Application.InchesToPoints(1.9)
                shp.Width = Application.InchesToPoints(2.6)
                ReScale shp
                
            Case "LB2"
                shp.Height = Application.InchesToPoints(3.1)
                shp.Width = Application.InchesToPoints(2.6)
                ReScale shp
                
            Case "LB1"
                shp.Height = Application.InchesToPoints(6.4)
                shp.Width = Application.InchesToPoints(3.6)
                ReScale shp
                
        End Select
    Next shp
    
    'Various worksheets
    For Each wks In ActiveWorkbook.Worksheets
        Select Case Left(wks.Name, 3)
            
            Case "(D)", "Dec" 'covers Decon Template and any Decons
                For Each shp In wks.Shapes
                    If shp.Type = msoOLEControlObject Then
                        shp.Height = Application.InchesToPoints(0.5)
                        shp.Width = Application.InchesToPoints(1)
                        Set obj = wks.OLEObjects(shp.Name)
                        If TypeOf obj.Object Is CommandButton Then obj.Object.Font.Size = 12
                        ReScale shp
                    End If
                Next shp
                
            Case "(C)", "COD"
                For Each shp In wks.Shapes
                    shp.Height = Application.InchesToPoints(0.6)
                    shp.Width = Application.InchesToPoints(0.75)
                    Set obj = wks.OLEObjects(shp.Name)
                    If TypeOf obj.Object Is CommandButton Then obj.Object.Font.Size = 12
                    ReScale shp
                Next shp
                
            Case "(St", "Sta"
                For Each shp In wks.Shapes
                    If shp.Type = msoOLEControlObject Then
                        shp.Height = Application.InchesToPoints(0.6)
                        shp.Width = Application.InchesToPoints(1.2)
                        Set obj = wks.OLEObjects(shp.Name)
                        If TypeOf obj.Object Is CommandButton Then obj.Object.Font.Size = 12
                        ReScale shp
                    End If
                Next shp

            Case "(LR", "LR "
                For Each shp In wks.Shapes
                    shp.Height = Application.InchesToPoints(0.5)
                    shp.Width = Application.InchesToPoints(1.2)
                    Set obj = wks.OLEObjects(shp.Name)
                    If TypeOf obj.Object Is CommandButton Then obj.Object.Font.Size = 12
                    ReScale shp
                Next shp
                
            Case "(P)", "Pre"
                For Each shp In wks.Shapes
                    Select Case Left(shp.Name, 3)
                        Case "Imp" 'Import buttons
                            Set obj = wks.OLEObjects(shp.Name)
                            If TypeOf obj.Object Is CommandButton Then obj.Object.Font.Size = 12
                            shp.Height = Application.InchesToPoints(0.72)
                            shp.Width = Application.InchesToPoints(1.3)
                            ReScale shp
                            
                        Case "Cmd" 'Add/Remove prop buttons
                            Set obj = wks.OLEObjects(shp.Name)
                            If TypeOf obj.Object Is CommandButton Then obj.Object.Font.Size = 12
                            shp.Height = Application.InchesToPoints(0.72)
                            shp.Width = Application.InchesToPoints(1)
                            ReScale shp
                    End Select
                Next shp
                
                wks.Shapes("SpinButton_Cont").Height = Application.InchesToPoints(0.48)
                wks.Shapes("SpinButton_Cont").Width = Application.InchesToPoints(0.42)
                
        End Select
    Next wks
    
    Application.ScreenUpdating = True
    
End Sub

Sub ReScale(obj As Shape)
'When fixing buttons, they need to be scaled up and then back to normal to "refresh" them. Microsoft is dumb sometimes.

    With obj
        .ScaleHeight 1.25, msoFalse
        .ScaleWidth 1.25, msoFalse
        .ScaleHeight 0.8, msoFalse
        .ScaleWidth 0.8, msoFalse
    End With
    
End Sub
Sub UpdateFooters()

    Application.ScreenUpdating = False
    
    Dim DefaultSheets As New Collection
    
    DefaultSheets.Add "Pre-STRmix Template"
    DefaultSheets.Add "Decon Template"
    DefaultSheets.Add "LR Template"
    DefaultSheets.Add "CODIS Template"
    DefaultSheets.Add "Standards Template"
    DefaultSheets.Add "1P Template"
    
    Dim wSheet As Variant
    For Each wSheet In DefaultSheets
        With Sheets(wSheet).PageSetup
            .CenterFooter = "&B&10&""Perpetua""" & Sheets("STRlite Settings").Range("Lab_Name") & Chr(10) & _
                            "&B&10&""Perpetua""" & "STRlite v" & Sheets("STRlite Settings").Range("Version_STRlite") & _
                            " compatible with STRmix" & Chr(153) & " v" & Sheets("STRlite Settings").Range("Version_STRmix")
        End With
    Next wSheet
    
    '2P combo sheet is pretty tight with an LR, so we use the left & center footers:
    With Sheets("2P Template").PageSetup
        .LeftFooter = "&B&10&""Perpetua""" & Sheets("STRlite Settings").Range("Lab_Name")
        .CenterFooter = "&10&""Perpetua""" & "STRlite v" & Sheets("STRlite Settings").Range("Version_STRlite") & _
                        " compatible with STRmix" & Chr(153) & " v" & Sheets("STRlite Settings").Range("Version_STRmix")
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Sub ResetPrintAreas()

    Sheets("Master").PageSetup.PrintArea = "$A$1:$AB$37"
    Sheets("About").PageSetup.PrintArea = "$B$2:$O$32"
    Sheets("STRlite Settings").PageSetup.PrintArea = "$A$1:$G$119"
    
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        Select Case Left(ws.Name, 3)
            
            Case "(D)", "Dec" 'covers Decon Template and any Decons
                ws.PageSetup.PrintArea = "$C$2:$M$48"
                
            Case "(C)", "COD"
                ws.PageSetup.PrintArea = "$B$2:$K$45"
                
            Case "(St", "Sta"
                ws.PageSetup.PrintArea = "$B$2:$H$37"

            Case "(LR", "LR "
                ws.PageSetup.PrintArea = "$B$2:$P$95"
                
            Case "(P)", "Pre"
                ws.PageSetup.PrintArea = "$B$1:$N$34"
            
            Case "(1P", "1P "
                ws.PageSetup.PrintArea = "$B$2:$L$35"
            
            Case "(2P", "2P "
                ws.PageSetup.PrintArea = "$B$2:$L$35"
                
        End Select
    
'        ws.PageSetup.Zoom = False
'        ws.PageSetup.FitToPagesTall = 1
'        ws.PageSetup.FitToPagesWide = 1
'        ws.PageSetup.CenterHorizontally = True
'        ws.PageSetup.CenterVertically = True

    Next ws

End Sub

Sub About()

' Generate random value between 1 and 15.
    Dim rndQuote As Integer
    rndQuote = CInt(Int((15 * Rnd()) + 1))

    Dim MainText As String
    Dim QuoteText As String
    Dim Version As String
    Dim VersionDate As String
    
    Version = "STRlite v2.1"
    VersionDate = "July 2020"
    
    MainText = CStr(Version & vbNewLine & VersionDate _
        & vbNewLine & vbNewLine & "by Melanie Wagner" _
        & vbNewLine & "Indiana State Police Laboratory" _
        & vbNewLine & vbNewLine)
    
    Select Case rndQuote
    
    Case 1
        QuoteText = _
        """The struggle itself is enough to fill a man's heart." & vbNewLine & _
        "One must imagine Sisyphus happy.""  -Albert Camus"
    
    Case 2
        QuoteText = _
        """Science moves with the spirit of an adventure" & vbNewLine & _
        "characterized both by youthful arrogance and by" & vbNewLine & _
        "the belief that truth, once found, would be simple" & vbNewLine & _
        "as well as pretty.""  -James Watson"
    
    Case 3
        QuoteText = _
        """The scientist only imposes two things, namely truth and" & vbNewLine & _
        "sincerity; imposes them upon himself and upon other scientists.""" & vbNewLine & _
        "-Erwin Schrodinger"
    
    Case 4
        QuoteText = _
        """If you wish to make an apple pie from scratch," & vbNewLine & _
        "you must first invent the universe.""  -Carl Sagan"
    
    Case 5
        QuoteText = _
        """To err is human, but to really foul things up" & vbNewLine & _
        "you need a computer.""  -Paul R. Ehrlich"
    
    Case 6
        QuoteText = _
        """To kill an error is as good a service as, and sometimes" & vbNewLine & _
        "even better than, the establishing of a new truth or fact.""" & vbNewLine & _
        "-Charles Darwin"
        
    Case 7
        QuoteText = _
        """To know that we know what we know, and to know that we" & vbNewLine & _
        "do not know what we do not know, that is true knowledge.""" & vbNewLine & _
        "-Nicolaus Copernicus"
    
    Case 8
        QuoteText = _
        """Science never solves a problem without" & vbNewLine & _
        "creating ten more."" -George Bernard Shaw"
    
    Case 9
        QuoteText = _
        """Science is a wonderful thing if one does not have" & vbNewLine & _
        "to earn one's living at it."" -Albert Einstein"
    
    Case 10
       QuoteText = _
        """If your experiment needs statistics, you ought to have" & vbNewLine & _
        "done a better experiment."" -Ernest Rutherford"
        
    Case 11
        QuoteText = _
        """The question is not whether machines think," & vbNewLine & _
        "but whether men do."" -B.F. Skinner"
    
    Case 12
        QuoteText = _
        """Any sufficiently advanced technology is" & vbNewLine & _
        "indistinguishable from magic."" -Arthur C. Clarke"
    
    Case 13
        QuoteText = _
        """We are stuck with technology, when what we really want" & vbNewLine & _
        "is just stuff that works."" -Douglas Adams"
    
    Case 14
        QuoteText = _
        """The universe is under no obligation to make sense to you.""" & vbNewLine & _
        "-Neil deGrasse Tyson"
    
    Case 15
        QuoteText = _
        """If you can't explain it simply, you don't understand" & vbNewLine & _
        "it well enough."" -Albert Einstein"
    
    End Select
    
    MsgBox MainText & QuoteText, 64, "About STRlite"

End Sub


Sub AddReferences(wbk As Workbook)

    AddRef wbk, "{000204EF-0000-0000-C000-000000000046}", "VBA"
    AddRef wbk, "{00020813-0000-0000-C000-000000000046}", "Excel"
    AddRef wbk, "{00020430-0000-0000-C000-000000000046}", "stdole"
    AddRef wbk, "{420B2830-E718-11CF-893D-00A0C9054228}", "Scripting"
    AddRef wbk, "{0D452EE1-E08F-101A-852E-02608C4D0BB4}", "MSForms"
    AddRef wbk, "{BEE4BFEC-6683-3E67-9167-3C0CBC68F40A}", "System"
    AddRef wbk, "{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", "VBScript_RegExp_55"
    AddRef wbk, "{F5078F18-C551-11D3-89B9-0000F81FE221}", "MSXML2"
    AddRef wbk, "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", "Office"
    AddRef wbk, "{215D64D2-031C-33C7-96E3-61794CD1EE61}", "System_Windows_Forms"
    
End Sub

Sub AddRef(wbk As Workbook, sGuid As String, sRefName As String)

    Dim i As Integer
    On Error GoTo EH
    With wbk.VBProject.References
        For i = 1 To .Count
            If .Item(i).Name = sRefName Then
               Exit For
            End If
        Next i
        If i > .Count Then
           .AddFromGuid sGuid, 0, 0 ' 0,0 should pick the latest version installed on the computer
        End If
    End With
EX: Exit Sub
EH: MsgBox "Error in 'AddRef'" & vbCrLf & vbCrLf & Err.Description
    Resume EX
    Resume ' debug code
End Sub


Public Sub DebugPrintExistingRefs()
    Dim i As Integer
    With Application.ThisWorkbook.VBProject.References
        For i = 1 To .Count
            Debug.Print "    AddRef wbk, """ & .Item(i).GUID & """, """ & .Item(i).Name & """"
        Next i
    End With
End Sub




