Attribute VB_Name = "RibbonAction"
'I did write the RibbonUI callback subs in this module.
Dim v As Variant
Option Explicit

'Callback for CaseDate onChange (updates case date)
Sub ChangeDate(control As IRibbonControl, Text As String)
    Range("CaseDate") = Text
End Sub

'Callback for CaseDate getText (shows current case date in entry box)
Sub gtCaseDate(control As IRibbonControl, ByRef Text)
    Text = Range("CaseDate")
End Sub

'Callback for Analyst onChange (updates analyst)
Sub ChangeAnalyst(control As IRibbonControl, Text As String)
    Range("Analyst") = Text
End Sub

'Callback for Analyst getText (shows current Analyst in entry box)
Sub gtAnalyst(control As IRibbonControl, ByRef Text)
    Text = Range("Analyst")
End Sub


'Callback for SortBy dropdown onAction
Sub ribSortBy(control As IRibbonControl, ByRef dropdownID As String, ByRef selectedIndex As Variant)
    Sheets("Master").Range("Dest_SortType").Value = IIf(selectedIndex = 0, "Type", "Case")
End Sub

'Callback for SortBy getSelectedItemID (default selection = Type and SortBy = true)
Sub gtSortBy(control As IRibbonControl, ByRef selectedIndex As Variant)
    selectedIndex = 0
End Sub


Sub ribRefreshSheets(control As IRibbonControl)
    Sheets("Master").RefreshMaster
    RibbonModule.RefreshRibbon
End Sub


Sub ribImportGMID(control As IRibbonControl)
    FormImportGenTable.Show vbModeless
End Sub

Sub ribImportDecon(control As IRibbonControl)
    Call Decon.ImportOneDecon
End Sub

Sub ribImportDeconMult(control As IRibbonControl)
    FormImportDecons.Show vbModeless
End Sub

Sub ribImportLR(control As IRibbonControl)
    FormImportLR.Show vbModeless
End Sub



'Callback for STRliteHome onAction
Sub ribHome(control As IRibbonControl)
    Sheets("Master").Select
    Call Sheets("Master").RefreshMaster
End Sub

'Callback for PreSTRmixList dynamic menu
Sub ribPreSTRmixList(control As IRibbonControl, ByRef strXML)
    If GMIDSheets Is Nothing Then Admin.LoadSheetLists
    strXML = xmlSheetList(GMIDSheets)
End Sub

'Callback for DeconList dynamic menu
Sub ribDeconList(control As IRibbonControl, ByRef strXML)
    If DeconSheets Is Nothing Then Admin.LoadSheetLists
    strXML = xmlSheetList(DeconSheets)
End Sub

'Callback for LRList dynamic menu
Sub ribLRList(control As IRibbonControl, ByRef strXML)
    If LRSheets Is Nothing Then Admin.LoadSheetLists
    strXML = xmlSheetList(LRSheets)
End Sub

'Callback for StandardsList dynamic menu
Sub ribStandardsList(control As IRibbonControl, ByRef strXML)
    If StandardSheets Is Nothing Then Admin.LoadSheetLists
    strXML = xmlSheetList(StandardSheets)
End Sub


'Callback for CODISList dynamic menu
Sub ribCODISList(control As IRibbonControl, ByRef strXML)
    If CODISSheets Is Nothing Then Admin.LoadSheetLists
    strXML = xmlSheetList(CODISSheets)
End Sub


'Callback for 1PList dynamic menu
Sub rib1PList(control As IRibbonControl, ByRef strXML)
    If SingleSheets Is Nothing Then Admin.LoadSheetLists
    strXML = xmlSheetList(SingleSheets)
End Sub

'Callback for 2PList dynamic menu
Sub rib2PList(control As IRibbonControl, ByRef strXML)
    If DoubleSheets Is Nothing Then Admin.LoadSheetLists
    strXML = xmlSheetList(DoubleSheets)
End Sub

'Callback for SettingsButton onAction
Sub ribSettings(control As IRibbonControl)
    Sheets("STRlite Settings").Visible = True
    Sheets("STRlite Settings").Select
End Sub

'Callback for AdminToolButton onAction
Sub ribAdminTools(control As IRibbonControl)
    FormAdmin.Show vbModeless
End Sub

'Callback for AboutButton onAction
Sub ribAbout(control As IRibbonControl)
    Call Admin.About
End Sub

'Callback for HelpButton onAction
Sub ribHelp(control As IRibbonControl)
    MsgBox "1) Remember to select folders, not files, when importing decons & LRs." & vbNewLine & vbNewLine & _
            "2) The case number fills on the combo sheets when you import a decon." & vbNewLine & vbNewLine & _
            "3) Don't use weird characters in your sample names. The following are ok: _ - + ( ) , . #", vbOKOnly + vbInformation, "Before you call Melanie..."
End Sub


Function xmlSheetList(Source As Dictionary) As String

    Dim q As String: q = Chr(34) ' " character
    Dim xmlID As String
    Dim xmlLabel As String
    Dim xmlTag As String
    Dim OnAction As String: OnAction = "onAction=" & q & "ribSelectSheet" & q
    Dim xmlLine As String: xmlLine = ""
    
    For Each v In Source.Keys
        'Ok. IRibbonControls are picky about the xml control IDs.
        'No spaces, parentheses, commas, pluses, and can't start with a number.
        'But nobody will ever see the control ID so it can look ridiculous as long as it stays unique.
        'Replaced the parentheses with "x" and "y" and replaced spaces with underscores. Also allowed for + and commas.
        xmlID = "id=" & q & Replace(Replace(Replace(Replace(Replace(Replace( _
            v, "(", "x", 1), ")", "y", 1), " ", "_"), "+", "p", 1), ",", "c", 1), "#", "n", 1) & q
        xmlLabel = "label=" & q & v & q
        xmlTag = "tag=" & q & v & q 'the tag is where we store the real worksheet name to grab from ribSelectSheet
        xmlLine = xmlLine & "<button " & xmlID & " " & xmlLabel & " " & xmlTag & " " & OnAction & "/>" & vbNewLine
    Next v
        
    'Debug.Print xmlLine
    
    'Apparently the xmlns is required. Seems redundant because it's already declared default at the root level, but whatever.
    xmlSheetList = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbNewLine & xmlLine & vbNewLine & "</menu>"

    'Debug.Print xmlSheetList
    
End Function


Sub ribSelectSheet(control As IRibbonControl)
    Application.ScreenUpdating = False
    Sheets(control.Tag).Select
    Application.ScreenUpdating = True
End Sub
