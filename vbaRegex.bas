Attribute VB_Name = "vbaRegex"
' Instructions:
' You need to enable this to be able to use the functions in this module.
' Go to Tools -> References, find "Microsoft VBScript Regular Expressions 5.5" and enable it.

' Finds the matching regexstring "strPattern" in "myRange" and returns it.
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
            regexMatch = "no match for: " & strPattern
        End If
    End If
End Function

' Finds the matching regexstring "strPattern" in "myRange", and replaces it with the provided string "replaceWith"
' returns the modified String
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
' TODO: Change so we can pass args to this function to 
' configure some settings, in the mean time, change below 
' if you desire other functionality
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
