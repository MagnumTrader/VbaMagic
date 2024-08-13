Attribute VB_Name = "vbaRegex"
' Instructions:
' First you need to enable Regular expressions in Excel to be able to use the functions in this module.
' In the Visual Basic window go to Tools -> References, find "Microsoft VBScript Regular Expressions 5.5" and enable it.

' Will return the *first* match, matching the pattern "strPattern" in the string "myRange" and returns the match.
' usage: =regexMatch("string to be matched against"; [a-z]+), will return "string"
Function regexMatch(myRange As Range, strPattern As String) As String
    Dim regEx As RegExp
    Dim strInput As String
    Dim matches As Object

    If strPattern <> "" Then
        strInput = myRange.Value
        Set regEx = configureRegex(strPattern)
        Set matches = regEx.Execute(strInput)
        If matches.Count > 0 Then
            regexMatch = matches(0)
        Else
            regexMatch = "No match for: " & strPattern
        End If
    End If

End Function

' Finds the matching regexstring "strPattern" in "myRange", and replaces it with the provided string "replaceWith"
' returns the modified String
' usage: in a cell type =regexReplace("string to be matched against"; [a-z]+, "better string") fill return "better string to be matched against"
' Global is set per default (in the configureRegex Function) to false, so will only replace the first match.
Function regexReplace(myRange As Range, strPattern As String, replaceWith As String) As String
    Dim regEx As RegExp
    Dim strInput As String

    If strPattern <> "" Then
        strInput = myRange.Value
        Set regEx = configureRegex(strPattern)
        regexReplace = regEx.Replace(strInput, replaceWith)
    Else
        regexReplace = strInput
    End If

End Function

' Global configuration for Regex
Private Function configureRegex(strPattern As String)
    Dim regEx As New RegExp

    With regEx
            .Global = False
            .MultiLine = False
            .IgnoreCase = False
            .Pattern = strPattern
    End With

    Set configureRegex = regEx

End Function
