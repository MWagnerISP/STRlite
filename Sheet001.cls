VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim v As Variant, counter As Integer, i As Integer


'****************************************************
'                Command Buttons
'****************************************************

Private Sub cmdDelete1PPage_Click()

    If Not LBselected(Me.LB2_1P) Then Exit Sub
    
    answer = MsgBox("Are you sure you want to permanently" & vbNewLine & "delete the selected Single Source sheet(s)?" & _
    vbNewLine & vbNewLine & "You cannot undo this.", vbCritical + vbYesNo, "Delete Single Source(s)?")
    
    If answer = vbNo Or answer = vbCancel Then Exit Sub
    
    DeletePage LB2_1P, SingleSheets
    
End Sub

Private Sub cmdDelete2PPage_Click()

    If Not LBselected(Me.LB2_2P) Then Exit Sub
    
    answer = MsgBox("Are you sure you want to permanently" & vbNewLine & "delete the selected 2P sheet(s)?" & _
    vbNewLine & vbNewLine & "You cannot undo this.", vbCritical + vbYesNo, "Delete 2P Worksheet(s)?")
    
    If answer = vbNo Or answer = vbCancel Then Exit Sub
    
    DeletePage LB2_2P, DoubleSheets

End Sub

Private Sub cmdDeleteAll_Click()

    If Not LBselected(Me.LB1_AllPages) Then Exit Sub
    
    answer = MsgBox("Are you sure you want to permanently" & vbNewLine & "delete the selected sheet(s)?" & _
    vbNewLine & vbNewLine & "You cannot undo this.", vbCritical + vbYesNo, "Delete Worksheet(s)?")
    
    If answer = vbNo Or answer = vbCancel Then Exit Sub
    
    counter = 0
    Dim wsSelection As String
    
    Application.DisplayAlerts = False
    
    For i = 0 To LB1_AllPages.ListCount - 1
        If LB1_AllPages.Selected(i - counter) Then
            wsSelection = LB1_AllPages.List(i - counter)
            Worksheets(wsSelection).Delete
            LB1_AllPages.RemoveItem (i - counter)
            If AllSheets.Exists(wsSelection) Then AllSheets.Remove (wsSelection)
            
            Select Case Left(wsSelection, 3) 'figure out which dictionary this worksheet belonged to, and remove it
                Case "(P)": If GMIDSheets.Exists(wsSelection) Then GMIDSheets.Remove (wsSelection)
                Case "(D)": If DeconSheets.Exists(wsSelection) Then DeconSheets.Remove (wsSelection)
                Case "(1P": If SingleSheets.Exists(wsSelection) Then SingleSheets.Remove (wsSelection)
                Case "(2P": If DoubleSheets.Exists(wsSelection) Then DoubleSheets.Remove (wsSelection)
                Case "(LR": If LRSheets.Exists(wsSelection) Then LRSheets.Remove (wsSelection)
                Case "(C)": If CODISSheets.Exists(wsSelection) Then CODISSheets.Remove (wsSelection)
                Case "(St": If StandardSheets.Exists(wsSelection) Then StandardSheets.Remove (wsSelection)
            End Select
            
            counter = counter + 1
            
        End If
    Next i

    With Me
        .LB1_AllPages.List = Admin.SortAllSheets(Me.Range("Dest_SortType").Value).Keys
        .LB2_PreSTRmix.List = PubFun.SortDictionaryByKey(GMIDSheets).Keys
        .LB2_Decon.List = PubFun.SortDictionaryByKey(DeconSheets).Keys
        .LB_LRs.List = PubFun.SortDictionaryByKey(LRSheets).Keys
        .LB_Stds.List = PubFun.SortDictionaryByKey(StandardSheets).Keys
        .LB_CODIS.List = PubFun.SortDictionaryByKey(CODISSheets).Keys
        .LB2_1P.List = PubFun.SortDictionaryByKey(SingleSheets).Keys
        .LB2_2P.List = PubFun.SortDictionaryByKey(DoubleSheets).Keys
    End With
    
    Application.DisplayAlerts = True
    
End Sub

Private Sub cmdDeleteCODISPage_Click()

    If Not LBselected(Me.LB_CODIS) Then Exit Sub
    
    answer = MsgBox("Are you sure you want to permanently" & vbNewLine & "delete the selected CODIS sheet(s)?" & _
    vbNewLine & vbNewLine & "You cannot undo this.", vbCritical + vbYesNo, "Delete CODIS?")
    
    If answer = vbNo Or answer = vbCancel Then Exit Sub
    
    DeletePage LB_CODIS, CODISSheets

End Sub

Private Sub cmdDeleteDecon_Click()

    If Not LBselected(Me.LB2_Decon) Then Exit Sub
    
    answer = MsgBox("Are you sure you want to permanently" & vbNewLine & "delete the selected decon(s)?" & _
    vbNewLine & vbNewLine & "You cannot undo this.", vbCritical + vbYesNo, "Delete Deconvolution(s)?")
    
    If answer = vbNo Or answer = vbCancel Then Exit Sub
    
    DeletePage LB2_Decon, DeconSheets
    
End Sub

Private Sub cmdDeleteLRPage_Click()
    
    If Not LBselected(Me.LB_LRs) Then Exit Sub
    
    If LRSheets Is Nothing Then Set LRSheets = New Scripting.Dictionary
    
    answer = MsgBox("Are you sure you want to permanently" & vbNewLine & "delete the selected LR sheet(s)?" & vbNewLine & vbNewLine & _
    "You cannot undo this.", vbCritical + vbYesNo, "Delete LR Page(s)?")
    
    If answer = vbNo Or answer = vbCancel Then Exit Sub
    
    DeletePage LB_LRs, LRSheets
    
End Sub

Private Sub cmdDeletePreSTRmix_Click()

    If Not LBselected(Me.LB2_PreSTRmix) Then Exit Sub
    
    answer = MsgBox("Are you sure you want to permanently" & vbNewLine & "delete the selected Pre-STRmix sheet(s)?" & vbNewLine & vbNewLine & _
    "You cannot undo this.", vbCritical + vbYesNo, "Delete Pre-STRmix Page(s)?")
    
    If answer = vbNo Or answer = vbCancel Then Exit Sub
    
    DeletePage LB2_PreSTRmix, GMIDSheets
    
End Sub

Private Sub cmdDeleteStdPage_Click()

    If Not LBselected(Me.LB_Stds) Then Exit Sub
    
    answer = MsgBox("Are you sure you want to permanently" & vbNewLine & "delete the selected standard sheet(s)?" & _
    vbNewLine & vbNewLine & "You cannot undo this.", vbCritical + vbYesNo, "Delete Standard Page(s)?")
    
    If answer = vbNo Or answer = vbCancel Then Exit Sub
    
    DeletePage LB_Stds, StandardSheets
    
End Sub

Private Sub cmdImport1Decon_Click()
'Prompts for a single decon folder

    Application.ScreenUpdating = False
    Decon.ImportOneDecon
    Application.ScreenUpdating = True
End Sub

Private Sub cmdImportDecons_Click()

    Application.ScreenUpdating = False
    FormImportDecons.Show
    Application.ScreenUpdating = True
End Sub

Private Sub cmdImportGenTable_Click()
    Application.ScreenUpdating = False
    FormImportGenTable.Show
    Application.ScreenUpdating = True
End Sub

Private Sub cmdImportLRs_Click()
    Application.ScreenUpdating = False
    FormImportLR.Show
    Application.ScreenUpdating = True
End Sub

Private Sub cmdImportStandards_Click()
    Application.ScreenUpdating = False
    FormImportGenTable.Show
    Application.ScreenUpdating = True
End Sub


Private Sub cmdPrintSTRlite_Click()

    If Me.LB1_AllPages.ListIndex = -1 Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim printSelected As New Scripting.Dictionary
    Dim lastSheet As Worksheet: Set lastSheet = Worksheets("Master")
    Dim arrPrint() As String, strPrintSheets As String
    
    
    For i = 0 To LB1_AllPages.ListCount - 1
        If LB1_AllPages.Selected(i) Then
            printSelected.Add LB1_AllPages.List(i), ""
            'Re-order worksheets so they print in the selected order
            Worksheets(CStr(LB1_AllPages.List(i))).Move After:=lastSheet
            Set lastSheet = Worksheets(CStr(LB1_AllPages.List(i)))
        End If
    Next i
    
    'Turn the dictionary keys into a STRING array (as opposed to variant array) by joining/splitting
    strPrintSheets = Join(printSelected.Keys(), ",")
    arrPrint = Split(strPrintSheets, ",")
    
    Worksheets(arrPrint).PrintPreview

    Me.Select
    
    Application.ScreenUpdating = True

End Sub

'****************************************************
'                List Boxes
'****************************************************

Private Sub LB1_AllPages_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    SelectPage LB1_AllPages

End Sub


Private Sub LB2_PreSTRmix_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.LB2_PreSTRmix.ListIndex = -1 Then Exit Sub

    Worksheets(Me.LB2_PreSTRmix.List(Me.LB2_PreSTRmix.ListIndex)).Select
    
End Sub

Private Sub LB2_Decon_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.LB2_Decon.ListIndex = -1 Then Exit Sub

    Worksheets(Me.LB2_Decon.List(Me.LB2_Decon.ListIndex)).Select

End Sub

Private Sub LB_LRs_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.LB_LRs.ListIndex = -1 Then Exit Sub

    Worksheets(Me.LB_LRs.List(Me.LB_LRs.ListIndex)).Select
    
End Sub

Private Sub LB_CODIS_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.LB_CODIS.ListIndex = -1 Then Exit Sub

    Worksheets(Me.LB_CODIS.List(Me.LB_CODIS.ListIndex)).Select
    
End Sub

Private Sub LB_Stds_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.LB_Stds.ListIndex = -1 Then Exit Sub

    Worksheets(Me.LB_Stds.List(Me.LB_Stds.ListIndex)).Select
    
End Sub

Private Sub LB2_1P_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.LB2_1P.ListIndex = -1 Then Exit Sub

    Worksheets(Me.LB2_1P.List(Me.LB2_1P.ListIndex)).Select
    
End Sub

Private Sub LB2_2P_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.LB2_2P.ListIndex = -1 Then Exit Sub

    Worksheets(Me.LB2_2P.List(Me.LB2_2P.ListIndex)).Select
    
End Sub



'****************************************************
'                General
'****************************************************
Function LBselected(LB As MSForms.ListBox) As Boolean

    Dim i As Integer

    With LB
        For i = 0 To .ListCount - 1
            LBselected = LBselected Or .Selected(i)
        Next i
    End With
    
End Function

Private Sub Worksheet_Activate()

    RefreshMaster
    
    RibbonModule.RefreshRibbon

End Sub

Sub SelectPage(LB As MSForms.ListBox)

    If LB.ListIndex = -1 Then Exit Sub
    If Not WorksheetExists(LB.List(LB.ListIndex)) Then Exit Sub
    
    Worksheets(LB.List(LB.ListIndex)).Select
    ActiveWindow.ScrollRow = 1

End Sub

Sub DeletePage(LB As MSForms.ListBox, dictSheets As Dictionary)

    counter = 0
    Dim wsSelection As String
    
    Application.DisplayAlerts = False
    
    For i = 0 To LB.ListCount - 1
        If LB.Selected(i - counter) Then
            wsSelection = LB.List(i - counter)
            Worksheets(wsSelection).Delete
            LB.RemoveItem (i - counter)
            If AllSheets.Exists(wsSelection) Then AllSheets.Remove (wsSelection)
            If dictSheets.Exists(wsSelection) Then dictSheets.Remove (wsSelection)
            counter = counter + 1
        End If
    Next i

    LB.List = dictSheets.Keys
    LB1_AllPages.List = AllSheets.Keys 'always refresh AllSheets LB
    
    Application.DisplayAlerts = True

End Sub

Public Sub RefreshMaster()

    LoadSheetLists

    With Me
        .LB1_AllPages.List = Admin.SortAllSheets(Me.Range("Dest_SortType").Value).Keys
        .LB2_PreSTRmix.List = PubFun.SortDictionaryByKey(GMIDSheets).Keys
        .LB2_Decon.List = PubFun.SortDictionaryByKey(DeconSheets).Keys
        .LB_LRs.List = PubFun.SortDictionaryByKey(LRSheets).Keys
        .LB_Stds.List = PubFun.SortDictionaryByKey(StandardSheets).Keys
        .LB_CODIS.List = PubFun.SortDictionaryByKey(CODISSheets).Keys
        .LB2_1P.List = PubFun.SortDictionaryByKey(SingleSheets).Keys
        .LB2_2P.List = PubFun.SortDictionaryByKey(DoubleSheets).Keys
    End With
    
End Sub


Private Sub Worksheet_BeforeDelete()

    ThisWorkbook.Protect

    MsgBox "Why would you delete the Master tab??", vbCritical + vbOKOnly, "You're why I lock these things down..."

End Sub


Private Sub Worksheet_Deactivate()
    ThisWorkbook.Unprotect
End Sub
