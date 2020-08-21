VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet300"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim LRCount As Integer
Dim button As Object
Dim LastUsedLR As Integer

'These represent the upper left corner ranges of each LR.
'Less error-prone than named ranges when duplicating sheets.

Const rngLR1 As String = "D11"
Const rngLR2 As String = "D24"
Const rngLR3 As String = "D37"
Const rngLR4 As String = "D58"
Const rngLR5 As String = "D71"
Const rngLR6 As String = "D84"

Private Function CountLastUsedLR() As Integer

    LastUsedLR = 1 'at least keep one out

    If Me.Range(rngLR2).Offset(0, 2).Value <> "" Then LastUsedLR = 2
    If Me.Range(rngLR3).Offset(0, 2).Value <> "" Then LastUsedLR = 3
    If Me.Range(rngLR4).Offset(0, 2).Value <> "" Then LastUsedLR = 4
    If Me.Range(rngLR5).Offset(0, 2).Value <> "" Then LastUsedLR = 5
    If Me.Range(rngLR6).Offset(0, 2).Value <> "" Then LastUsedLR = 6

End Function


Sub AddLR()
    
Application.ScreenUpdating = False

    Dim CurrentLRCount As Integer: CurrentLRCount = Me.Range("LR_Count").Value
    
    LRCount = IIf(CurrentLRCount >= 6, 6, CurrentLRCount + 1) 'can't have >6 LRs

    Call LRRows(LRCount)
    
    Me.Range("LR_Count").Value = LRCount

Application.ScreenUpdating = True

End Sub


Sub RemoveLR()

Application.ScreenUpdating = False

    Dim CurrentLRCount As Integer: CurrentLRCount = Me.Range("LR_Count").Value
    
    LRCount = IIf(CurrentLRCount <= 1, 1, CurrentLRCount - 1) 'can't have <1 LR
    
    Call LRRows(LRCount)
    
    Me.Range("LR_Count").Value = LRCount

Application.ScreenUpdating = True

End Sub


Sub LRRows(ByVal LRCount As Integer)

Dim EndOfHeader1 As Integer: EndOfHeader1 = 9 'Last row in page 1 header
Dim EndOfHeader2 As Integer: EndOfHeader2 = 56 'Last row in page 2 header
Dim StartOfPage2 As Integer: StartOfPage2 = 49 'First row in page 2 (header)
Dim VeryFirstRow As Integer: VeryFirstRow = 2
Dim VeryLastRow As Integer: VeryLastRow = 95
Dim LRsOnFirstPage As Integer: LRsOnFirstPage = 3
Dim LRSectionRows As Integer: LRSectionRows = 12 'Number of rows in each LR section (not including borders)


Select Case LRCount

Case Is = 6
    'Unhide all rows
        Me.Rows.Hidden = False
    
    'All buttons visible
        With Me
            For Each button In .OLEObjects
                button.Visible = True
            Next
        End With
    
    'Set print area to two pages
        Me.PageSetup.PrintArea = "$B$" & VeryFirstRow & ":$P$" & VeryLastRow


Case Is <= LRsOnFirstPage '(LRs 1-3)
    
    'Unhide/hide LRs 1-4 according to LRCount
    '(This way is faster than the For/Each loop:)
        'Hide all LRs first:
            Me.Rows((EndOfHeader1 + 1) & ":" & VeryLastRow).Hidden = True
        'Unhide LRs according to LRCount:
            Me.Rows((EndOfHeader1 + 1) & ":" & EndOfHeader1 + (LRCount * (LRSectionRows + 1) - 1)).Hidden = False
        'Unhide last row of page 1:
            Me.Rows(StartOfPage2 - 1).Hidden = False
    
    'Set print area to just one page
        Me.PageSetup.PrintArea = "$B$" & VeryFirstRow & ":$P$" & (StartOfPage2 - 1)
        

Case Is > LRsOnFirstPage '(LRs 4-6)

    'Unhide/hide props 5-8 according to LRCount
    '(This way is faster than the For/Each loop:)
    
        'Hide all LRs first:
            Me.Rows((EndOfHeader1 + 1) & ":" & VeryLastRow).Hidden = True
        'Unhide 2nd page header & LRs 1-4:
            Me.Rows(VeryFirstRow & ":" & EndOfHeader2).Hidden = False
        'Unhide 2nd page LRs according to LRCount:
            Me.Rows((EndOfHeader2 + 1) & ":" & EndOfHeader2 + ((LRCount - LRsOnFirstPage) * (LRSectionRows + 1) - 1)).Hidden = False
        'Unhide last row of page 2:
            Me.Rows(VeryLastRow).Hidden = False
    
    'Set print area to two pages
        Me.PageSetup.PrintArea = "$B$" & VeryFirstRow & ":$P$" & VeryLastRow


End Select


    'Hide/Unhide button sets
    For Each button In Me.OLEObjects
            'Hide all buttons whose names end in numbers < LRCount
            If Val(Right(button.Name, 1)) <= LRCount Then
                button.Visible = True
            Else
                button.Visible = False
            End If
    Next


    
End Sub


Private Sub Clear_LR1_Click()
    Call LR.ClearLR(Me.Range(rngLR1))
End Sub

Private Sub Clear_LR2_Click()
    Call LR.ClearLR(Me.Range(rngLR2))
End Sub

Private Sub Clear_LR3_Click()
    Call LR.ClearLR(Me.Range(rngLR3))
End Sub

Private Sub Clear_LR4_Click()
    Call LR.ClearLR(Me.Range(rngLR4))
    Me.Activate
End Sub

Private Sub Clear_LR5_Click()
    Call LR.ClearLR(Me.Range(rngLR5))
    Me.Activate
End Sub

Private Sub Clear_LR6_Click()
    Call LR.ClearLR(Me.Range(rngLR6))
    Me.Activate
End Sub


Private Sub cmdAddLR_Click()

    Application.ScreenUpdating = False
    Call AddLR
    ActiveWindow.ScrollRow = 1
    Application.ScreenUpdating = True
    
End Sub

Private Sub cmdRemoveLR_Click()

    Application.ScreenUpdating = False
    Call RemoveLR
    ActiveWindow.ScrollRow = 1
    Application.ScreenUpdating = True
    
End Sub

Private Sub CmdLR_ClearLRs_Click()
    
    Dim Choice As Variant
    Choice = MsgBox("Clear ALL likelihood ratios on this sheet?", vbYesNo + vbDefaultButton1 + vbQuestion, "Clear All LRs?")
    
    If Choice = vbNo Or Choice = vbCancel Then Exit Sub
    If Choice = vbYes Then
    
        Call LR.ClearLR(Me.Range(rngLR1))
        Call LR.ClearLR(Me.Range(rngLR2))
        Call LR.ClearLR(Me.Range(rngLR3))
        Call LR.ClearLR(Me.Range(rngLR4))
        Call LR.ClearLR(Me.Range(rngLR5))
        Call LR.ClearLR(Me.Range(rngLR6))
    
    End If
        
    Me.Activate
    
End Sub


Private Sub Import_LR1_Click()
    Call LR.Import_LRPrev(Me.Range(rngLR1))
End Sub

Private Sub Import_LR2_Click()
    Call LR.Import_LRPrev(Me.Range(rngLR2))
End Sub

Private Sub Import_LR3_Click()
    Call LR.Import_LRPrev(Me.Range(rngLR3))
End Sub

Private Sub Import_LR4_Click()
    Call LR.Import_LRPrev(Me.Range(rngLR4))
End Sub

Private Sub Import_LR5_Click()
    Call LR.Import_LRPrev(Me.Range(rngLR5))
End Sub

Private Sub Import_LR6_Click()
    Call LR.Import_LRPrev(Me.Range(rngLR6))
End Sub

Private Sub Worksheet_Activate()

    Application.ScreenUpdating = False
    
    Call CountLastUsedLR
    Call LRRows(LastUsedLR)
    Me.Range("LR_Count").Value = LastUsedLR
    ActiveWindow.ScrollRow = 1
    
    Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_BeforeDelete()

    If AllSheets Is Nothing Or LRSheets Is Nothing Then Admin.LoadSheetLists
    
    If AllSheets.Exists(Me.Name) Then AllSheets.Remove (Me.Name)
    If LRSheets.Exists(Me.Name) Then LRSheets.Remove (Me.Name)
    
End Sub
