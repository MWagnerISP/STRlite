Attribute VB_Name = "PubFun"
'This module is full of PUBlic FUNctions and subs but I called it PubFun because I'm twelve
'Does not include the object creation functions which are in the Factory module
Option Explicit

Public Function FixWorksheetName(inputSheetName As String) As String
'Worksheet names can't have certain characters:  \ / * ? : [ ]
'This intercepts sample names with illegal characters when naming sheets

    Dim tempName As String: tempName = inputSheetName
    
    'Replace illegal characters with another specific character:
    tempName = Replace(tempName, "[", "(", 1, -1, vbBinaryCompare)
    tempName = Replace(tempName, "]", ")", 1, -1, vbBinaryCompare)
    tempName = Replace(tempName, "\", "-", 1, -1, vbBinaryCompare)
    tempName = Replace(tempName, "/", "-", 1, -1, vbBinaryCompare)
    tempName = Replace(tempName, ":", "-", 1, -1, vbBinaryCompare)
    tempName = Replace(tempName, "*", "_", 1, -1, vbBinaryCompare)
    tempName = Replace(tempName, "?", "_", 1, -1, vbBinaryCompare)
    
'    Replace all illegal characters with a "_":
'    Dim i As Integer
'    Dim illegal(1 To 7) As String
'        illegal(1) = "\" 'chr(92)
'        illegal(2) = "/" 'chr(47)
'        illegal(3) = "*" 'chr(42)
'        illegal(4) = "?" 'chr(63)
'        illegal(5) = ":" 'chr(58)
'        illegal(6) = "[" 'chr(91)
'        illegal(7) = "]" 'chr(93)
'
'    For i = 1 To 7
'        tempName = Replace(tempName, illegal(i), "_", 1, -1, vbBinaryCompare)
'    Next i

    FixWorksheetName = tempName
    
End Function

Public Function WorksheetExists(sheetName As String, Optional wb As Workbook) As Boolean
    
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    On Error Resume Next
        Set sht = wb.Sheets(sheetName)
    On Error GoTo 0
    
    WorksheetExists = Not sht Is Nothing
    
End Function

Public Function SortObjectsByProperty(Source As Dictionary, property As String, Optional sortorder As XlSortOrder = xlAscending) As Object
'Melanie adapted this from the sub below and couldn't find anything else like it on the internet!! Super proud of this one!
'Source must be a uniform dictionary of objects that all contain the property in question

    Dim arrList As Object: Set arrList = CreateObject("System.Collections.ArrayList")
    Dim k As Variant, v As Variant
    
    'Put the property values in arrList
    For Each k In Source.Keys
        'CallByName= built-in VBA function, allows you to use "property" as a variable to return an actual object property.
        arrList.Add CallByName(Source(k), property, VbGet)
    Next k
    
    'Sort or reverse-sort
    arrList.Sort
    If sortorder = xlDescending Then arrList.Reverse
    
    'Create new dictionary for sorted list
    Dim dictNew As Dictionary: Set dictNew = New Scripting.Dictionary
    
    'Read through sorted arrList in order (k)
    For Each k In arrList
        'For each arrList item (k), find the Source key (v) that has a matching property
        For Each v In Source.Keys 'the keys in Source (v) are unique
            If Not dictNew.Exists(v) Then 'skip each v that's already been added
                If CallByName(Source(v), property, VbGet) = k Then
                    dictNew.Add v, Source(v)
                    'Debug.Print k, v
                End If
            End If
        Next v
    Next k
    
    Set SortObjectsByProperty = dictNew

End Function


Public Function SortDictionaryByKey(dict As Object, Optional sortorder As XlSortOrder = xlAscending) As Object
'This is some of the only code in STRlite that I didn't write. https://excelmacromastery.com/vba-dictionary/
    
    Dim arrList As Object
    Set arrList = CreateObject("System.Collections.ArrayList")
    
    ' Put keys in an ArrayList
    Dim key As Variant
    For Each key In dict
        arrList.Add key
    Next key
    
    ' Sort the keys
    arrList.Sort
    
    ' For descending order, reverse
    If sortorder = xlDescending Then
        arrList.Reverse
    End If
    
    ' Create new dictionary
    Dim dictNew As Object
    Set dictNew = CreateObject("Scripting.Dictionary")
    
    ' Read through the sorted keys and add to new dictionary
    For Each key In arrList
        dictNew.Add key, dict(key)
    Next key
    
    ' Clean up
    Set arrList = Nothing
    Set dict = Nothing
    
    ' Return the new dictionary
    Set SortDictionaryByKey = dictNew
        
End Function

Public Function PutLociInOrder(dict As Object) As Object
'Takes dict dictionary as input and returns a dictionary with the loci from dict in LociOrder order
    
    Dim v As Variant
    Dim dictNew As New Scripting.Dictionary
    
    For Each v In LociOrder.Keys
        If dict.Exists(v) Then dictNew.Add v, dict(v)
    Next v
    
    Set PutLociInOrder = dictNew
          
End Function

Public Function FindCell(FindWithin As Range, FindWhat As String, After As Range) As Range

    Set FindCell = FindWithin.Find _
    (What:=FindWhat, After:=After, _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

End Function


Public Sub ReplaceText(Target As Range, Text As String, Replacement As String)

    Target.Replace What:=Text, Replacement:=Replacement, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

End Sub


Public Function AlleleToDecimal(allele As Double, RepeatSize As Integer) As Double
' Converts partial variant alleles (e.g. 17.3) into a decimal/fraction of the whole repeat
' so that the stutter filter can tell what n-2 stutter is

    Dim strAllele As String
    Dim arrAllele() As String
    Dim whole As Double, partial As Double
    
    strAllele = CStr(allele)
    
    If InStr(1, strAllele, ".") > 0 Then
        arrAllele = Split(strAllele, ".")
        whole = CDbl(arrAllele(0))
        partial = CDbl(arrAllele(1)) / RepeatSize
        AlleleToDecimal = whole + partial
    Else:
        AlleleToDecimal = allele
    End If
        
End Function

Public Function DecimalToAllele(dec As Double, RepeatSize As Integer) As Double

    Dim strAllele As String
    Dim arrAllele() As String
    Dim whole As Double, partial As Double
    
    strAllele = CStr(dec)
    
    If InStr(1, strAllele, ".") > 0 Then
        arrAllele = Split(strAllele, ".")
        whole = arrAllele(0)
        partial = RepeatSize * CDbl("0." & arrAllele(1)) / 10
        DecimalToAllele = whole + partial
    Else:
        DecimalToAllele = dec
    End If

End Function


Public Function IsAlphaNumeric(char As String) As Boolean

    Dim upC As String: upC = UCase(char) 'UCase only affects letters
    
    'Uppercase alphabet is ASCII/Unicode 65-90
    IsAlphaNumeric = (AscW(upC) >= 65 And AscW(upC) <= 90) Or (VBA.IsNumeric(char))
    
End Function

Public Function CaseNumberMask(CaseInput As String, Optional Dash As Boolean, Optional OmitLeadingZeroes As Boolean, Optional Digits As Integer) As String
'If you're not ISP, you won't be using this, but you could maybe figure out how to adapt it to your own lab...
'RegEx is ridiculous. Just Google it and good luck.

    Dim regEx As New VBScript_RegExp_55.RegExp
    Dim casePattern As String
    
    'Indiana State Police case number RegEx pattern:
    casePattern = "(^[0-9]{2})([^OQYZ0-9])([\-]?)([0-9]{1,5}$)"
                'must begin with two numbers
                'followed by one letter but not O,Q,Y,or Z or numbers (set IgnoreCase = True)
                'may or may not have a dash
                'followed by 1 to 5 final numbers

    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = casePattern
    End With

    If regEx.test(CaseInput) Then
        'Here we get the four parts of a case number
        Dim yr As String: yr = regEx.Replace(CaseInput, "$1")
        Dim letter As String: letter = regEx.Replace(CaseInput, "$2")
        Dim inputDash As String: inputDash = regEx.Replace(CaseInput, "$3") 'we don't really care about the dash from the input, but it's part "$3"
        Dim lastNum As String: lastNum = regEx.Replace(CaseInput, "$4")
    Else:
        CaseNumberMask = "NotACase"
        Exit Function
    End If
    
    'Omit leading zeroes from lastNum
    If OmitLeadingZeroes Then
        Do While Left(lastNum, 1) = "0"
            lastNum = Replace(lastNum, "0", "", 1, 1)
        Loop
    End If
    
    'Set number of digits, if specified (this can basically cancel out OmitLeadingZeroes)
    Do While Len(lastNum) < Digits
        lastNum = "0" & lastNum
    Loop
    
    'Construct Case Number
    If Dash Then
        CaseNumberMask = yr & UCase(letter) & "-" & lastNum
    Else:
        CaseNumberMask = yr & UCase(letter) & lastNum
    End If
        
End Function

