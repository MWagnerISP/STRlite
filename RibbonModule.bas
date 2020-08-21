Attribute VB_Name = "RibbonModule"
'Ribbon icons by a graphic artist named Jojo Mendoza aka Deleket aka Hopstarter
'http://www.iconarchive.com/artist/hopstarter.html

Option Explicit
'I (Melanie) did not write the RibbonUI setup code in this module.

Public YourRibbon As IRibbonUI
Public MyTag As String

'This prevents the RibbonUI from getting "lost." The Ribbon object is saved as a number on worksheet STRlite Settings.
#If VBA7 Then
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef Source As Any, ByVal length As Long)
#Else
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef Source As Any, ByVal length As Long)
#End If

'Upon loading, the RibbonUI object is saved as a pointer (number)
Public Sub RibbonOnLoad(Ribbon As IRibbonUI)
   'Store pointer to IRibbonUI as "YourRibbon" and keep it in cell B14 of STRlite Settings:
    Set YourRibbon = Ribbon
    Sheets("STRlite Settings").Range("RibbonCell").Value = ObjPtr(Ribbon)
    
    'Make today the default date in the Ribbon
    Sheets("STRlite Settings").Range("CaseDate") = Date
    
    'Set default sort type for Master
    Sheets("Master").Range("Dest_SortType") = "Type"
    
End Sub

'Retrieves the RibbonUI object from the cell it was saved in

#If VBA7 Then
Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
#Else
Function GetRibbon(ByVal lRibbonPointer As Long) As Object
#End If
        Dim objRibbon As Object
        CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
        Set GetRibbon = objRibbon
        Set objRibbon = Nothing
End Function

'Refreshes the Ribbon (to be triggered every time a change is made that would affect the appearance of the buttons)
Sub RefreshRibbon()

    If YourRibbon Is Nothing Then
        'If the Ribbon is lost (is Nothing) then use function "GetRibbon" to retrieve it, then refresh:
        Set YourRibbon = GetRibbon(Sheets("STRlite Settings").Range("RibbonCell").Value)
        YourRibbon.Invalidate
    Else
        'If the Ribbon is not lost, just refresh (invalidate):
        YourRibbon.Invalidate
    End If
End Sub


