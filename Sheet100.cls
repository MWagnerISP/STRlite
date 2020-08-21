VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim PropCount As Integer
Dim Contributors As Integer

Private Sub Cmd_AddProp_Click()

    Call AddProp
    
End Sub

Private Sub Cmd_RemoveProp_Click()

    Call RemoveProp
    
End Sub

Private Sub cmdConvert_Click()
    
    Call Factory.ConvertToCombo(Me)
    
End Sub

Private Sub cmdHideCond_Click()

    Me.Range("E:F").EntireColumn.Hidden = True
    Me.cmdHideCond.Visible = False
    
End Sub

Private Sub ImportConditions_Click()

    Set FormConditionedGenotype.wsDest = Me 'pass this worksheet (Pre-STRmix) to the userform as a variable
    FormConditionedGenotype.Show
    
End Sub

Sub AddProp()

Application.ScreenUpdating = False

    Dim PropName As String
    
    Dim PropStart As Range
    Dim PropEnd As Range
    Dim PropRange As Range
    Dim PropDest As Range

    PropCount = Application.WorksheetFunction.CountA(Me.Range("Props")) 'Current number of propositions
    Contributors = Me.Range("Conts_Prop").Value
    
        If PropCount = 4 Or Contributors = 0 Or Contributors = 5 Then
            Exit Sub
        End If
        
    PropName = "Prop" & PropCount + 1 'Determines range name of next proposition
    
    With Me
        Set PropStart = .Range("Prop1").Offset(0, 1)
        Set PropEnd = .Range("Prop1").Offset(1, Contributors + 1)
        Set PropRange = .Range(PropStart.Address & ":" & PropEnd.Address)
    End With
        
    Me.Select
    PropRange.Select
    Selection.Copy
    Range(PropName).Offset(0, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
    Application.CutCopyMode = False
        
    With Range(PropName)
        .Value = PropCount + 1
        .Font.Name = "Perpetua"
        .Font.Size = "14"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Range(PropName).Resize(2, 1).Merge
        
    Range(PropName).Offset(0, 1).Value = "H1"
    Range(PropName).Offset(0, 1).Offset(1, 0).Value = "H2" 'Two offsets to accommodate merged PropName cell
    
    Range(PropName).Offset(0, 2).Select
    
Application.ScreenUpdating = True

End Sub

Sub RemoveProp()

Application.ScreenUpdating = False
    
    Dim PropName As String
    
    Dim PropStart As Range
    Dim PropEnd As Range
    Dim PropRange As Range
    Dim PropDest1 As String
    Dim PropDest2 As String

    PropCount = Application.WorksheetFunction.CountA(Me.Range("Props")) 'Current number of propositions
    Contributors = Range("Conts_Prop").Value
    
        If PropCount = 1 Or Contributors = 0 Or Contributors = 5 Then
            Exit Sub
        End If
    
    PropName = "Prop" & PropCount 'Determines range name of last proposition
    
    With Me
        Set PropStart = .Range(PropName).Offset(0, 1)
        Set PropEnd = .Range(PropName).Offset(1, 5)
        Set PropRange = .Range(PropStart.Address & ":" & PropEnd.Address)
    End With
        
    Me.Select
    PropRange.Select
    Selection.Clear
        'White background
        With Selection.Interior
            .PatternColorIndex = xlAutomatic
            .Color = 16777215
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
    
    PropDest1 = Range(PropName).Address
    PropDest2 = Range(PropName).Offset(1, 0).Address
    
    With Range(PropDest1, PropDest2)
        .ClearContents
        .Borders.LineStyle = xlNone
    End With
        
'    Range(PropName).Offset(0, 1).Clear
    'Range(PropName).Offset(0, 1).Offset(1, 0).Clear 'Two offsets to accommodate merged PropName cell
    
    Range(PropName).Offset(-3, 2).Select 'Select first cell of previous remaining proposition
    
Application.ScreenUpdating = True

End Sub


Sub ContributorChange(i As Integer, PropCount As Integer)

Application.ScreenUpdating = False

    Select Case i
    
        Case Is > 4
            Call ClearPropSpace
            Call FormatHeader(4)
    
        Case Is < 1
        
            Call ClearPropSpace
            Call FormatHeader(4)
            
        Case 1 To 4
        
            'Format header
            Call FormatHeader(i)
            
            'Determine unneeded contributor range and clear it
            With Range(Range("PropOrigin").Offset(1, 2 + i), Range("ContEnd"))
                .ClearContents
                .Borders.LineStyle = xlNone
            End With
            
            Call FormatFirstProp(i)
            
            'Copy format from first proposition to all propositions:
            Dim j As Integer
            For j = 2 To PropCount
            
                Me.Range("PropOrigin").Offset(1, 1).Resize(2, i + 1).Copy
                
                Range("Prop" & j).Offset(0, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
                With Range("Prop" & j)
                    .Value = j
                    .Font.Name = "Perpetua"
                    .Font.Size = "14"
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Resize(2, 1).Merge
                End With

            Next j
            
    End Select
    
        Me.Range("C3").Activate
        Application.CutCopyMode = False
        Application.ScreenUpdating = True

End Sub


Sub SpinButton_Cont_SpinDown()

    Call ChangeContributor

End Sub

Sub SpinButton_Cont_SpinUp()

    Call ChangeContributor
    
End Sub


Sub ChangeContributor()
'This handles input from spin button and intercepts the extremes

    If SpinButton_Cont.Value = 6 Then
        SpinButton_Cont.Value = 5
        Exit Sub
    End If
    
    If SpinButton_Cont.Value = -1 Then
        SpinButton_Cont.Value = 0
        Exit Sub
    End If
    
    Application.ScreenUpdating = False

    Dim i As Integer
    Dim Result As Variant

    i = SpinButton_Cont.Value
    
    Me.cmdConvert.Visible = IIf(i = 1 Or i = 2, True, False)
    
    If i > 4 Then
    
        Result = MsgBox("Determining more than 4 contributors will clear all" & vbNewLine & _
        "propositions and declare the sample inconclusive." & vbNewLine & vbNewLine & _
        "Proceed?", vbYesNo + vbDefaultButton1 + vbExclamation, "Inconclusive?")
    
        If Result = vbNo Then
            SpinButton_Cont.Value = 4
            Exit Sub
        End If
        
    End If
        
    If i < 1 Then
    
        Result = MsgBox("Cannot determine number of contributors?" & vbNewLine & vbNewLine & _
        "This will clear all propositions and declare the sample inconclusive." & vbNewLine & vbNewLine & _
        "Proceed?", vbYesNo + vbDefaultButton1 + vbExclamation, "Inconclusive?")
    
        If Result = vbNo Then
            SpinButton_Cont.Value = 1
            Exit Sub
        End If
        
    End If
        
    PropCount = Application.WorksheetFunction.CountA(Range("Props"))
    
    Call ContributorChange(i, PropCount)

End Sub

Sub ClearPropSpace()

    'Clear proposition space
    With Me.Range(Range("Prop1"), Me.Range("ContEnd"))
        .Clear
        With .Interior
            'White background
            .PatternColorIndex = xlAutomatic
            .Color = 16777215
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End With
    
End Sub


Sub FormatFirstProp(i As Integer)

    'Prop1 label + H1/H2:
        Range("PropOrigin").Offset(1, 1).Value = "H1"
        Range("PropOrigin").Offset(2, 1).Value = "H2"
        
        With Range("Prop1").Resize(2, 2)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            With .Font
                .Name = "Perpetua"
                .Size = 12
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Bold = True
                .Underline = xlUnderlineStyleNone
                .TintAndShade = 0
            End With
        End With
        
        With Range("Prop1")
            .Value = "1"
            .Font.Size = 14
        End With
        
        Range("Prop1").Resize(2, 1).Merge
    
    'Format first proposal table for i contributors:
    With Range("PropOrigin").Offset(1, 1).Resize(2, i + 1)

        With .Font
            .Name = "Perpetua"
            .Size = 12
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
        End With

        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 1
            .weight = xlThin
        End With
        
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 1
            .weight = xlThin
        End With
        
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ReadingOrder = xlContext
        .MergeCells = False
        .BorderAround weight:=xlMedium, ColorIndex:=0
        
    End With
    
    Range("PropOrigin").Offset(1, 1).Resize(2, 1).BorderAround weight:=xlMedium, ColorIndex:=0
    
    Range("PropOrigin").Offset(1, 2).Resize(2, 4).Locked = True 'Lock additional contributors
    Range("PropOrigin").Offset(1, 2).Resize(2, i).Locked = False 'Unlock current contributors (i)

End Sub

Sub FormatHeader(i As Integer)

        'Format header to i columns
        With Range("PropOrigin").Offset(0, 2).Resize(1, 4)
            .Borders.LineStyle = xlNone 'Take out old borders
            .UnMerge 'UnMerge Header
            With .Interior 'White background
                .PatternColorIndex = xlAutomatic
                .Color = 16777215
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        
        With Range("PropOrigin").Offset(0, 2).Resize(1, i)
            'Merge & center header
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
            .BorderAround weight:=xlMedium, ColorIndex:=0
        End With

End Sub

Private Sub Worksheet_BeforeDelete()

    If AllSheets Is Nothing Or GMIDSheets Is Nothing Then Admin.LoadSheetLists
    
    If AllSheets.Exists(Me.Name) Then AllSheets.Remove (Me.Name)
    If GMIDSheets.Exists(Me.Name) Then GMIDSheets.Remove (Me.Name)
    
End Sub
