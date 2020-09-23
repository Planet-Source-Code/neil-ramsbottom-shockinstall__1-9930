Attribute VB_Name = "modStringFuntions"

'All code contained within this module is FREEWARE and may be freely
'distributed and used within other applications providing that
'   i)No modifications are made to the original code
'
'  ii)No modifications are made to the comments
'
' iii)Credit is given in the applications about box or readme file.
'
'Neil Ramsbottom

Option Explicit

Public Function ReplaceCommandTag(strText As String, strTagName As String, strTagVal As String) As String

'Author:    Neil Ramsbottom
'Date:      29/01/2000
'Purpose:   Replaces the value of a command tag with a string.
'WWW:       http://www.nramsbottom.co.uk
'Example:   MsgBox (ReplaceCommandTag("The date today is %DATE%.","DATE",Format$(Date,"mm/dd/yyyy")))
'Notes:     This function is case sensitive, so many different values can be used, i.e. DATE, date, dATE, DaTE
'           I wrote this to help in the generation of 404 errors on a web server, where the Date and Time needed
'           to be added to a public string, before being sent to a browser.

'This module was written before I had VB6, but VB6 has the Replace() function

Dim i As Integer

Dim strTmpVal As String
Dim strTmpVal2 As String

If strText = "" Or Len(strText) = 0 Or strTagName = "" Or Len(strTagName) = 0 Or strTagVal = "" Or Len(strTagVal) = 0 Then
    ReplaceCommandTag = strText
    Exit Function
End If

For i = 1 To Len(strText)
    If Mid(strText, i, 1) = "%" Then
        'If the current character and strTagName and "%" =
        '"%" and strTagName &"%" then do . . .
        If Mid(strText, i, Len(strTagName & "%") + 1) = "%" & strTagName & "%" Then
            strTmpVal = Mid(strText, 1, i - 1) 'Store all of the string before the first %
            strTmpVal2 = Mid(strText, i + Len(strTagName) + 2) 'Store all of the string after the last "%"
            ReplaceCommandTag = strTmpVal & strTagVal & strTmpVal2 'Reconstruct the string and return value
            Exit Function 'Exits after first change
        End If
    End If
Next i

ReplaceCommandTag = strText

End Function

Public Function InvertSlashes(strText As String) As String

'Author:    Neil Ramsbottom
'Date:      28/01/2000
'Purpose:   Inverts any slashes in a string.
'Example:   Passing "c:\windows/desktop" will return "C:/windows\desktop"

'Pre-VB6 Function

Dim i As Integer

For i = 1 To Len(strText)

    If Mid(strText, i, 1) = "\" Then
        Mid(strText, i, 1) = "/"
    ElseIf Mid(strText, i, 1) = "/" Then
        Mid(strText, i, 1) = "\"
    End If

Next i

InvertSlashes = strText

End Function

Public Function InvertBackSlashes(strText As String) As String

'Author:    Neil Ramsbottom
'Date:      29/01/2000
'Purpose:   Inverts any backslashes within a string

'Pre-VB6 Function

Dim i As Integer

If strText = "" Or Len(strText) = 0 Then
    InvertBackSlashes = strText
    Exit Function
End If

For i = 1 To Len(strText)

    If Mid(strText, i, 1) = "\" Then
        Mid(strText, i, 1) = "/"
    End If
    
Next i

InvertBackSlashes = strText

End Function

Public Function InvertForwardSlashes(strText As String) As String

'Author:    Neil Ramsbottom
'Date:      29/01/2000
'Purpose:   Inverts any forward slashes within a string

'Pre-VB6 Function

Dim i As Integer

If strText = "" Or Len(strText) = 0 Then
    InvertForwardSlashes = strText
    Exit Function
End If

For i = 1 To Len(strText)

    If Mid(strText, i, 1) = "/" Then
        Mid(strText, i, 1) = "\"
    End If
    
Next i

InvertForwardSlashes = strText

End Function

Public Function WindowsDirectory() As String

'Author:    Neil Ramsbottom
'Date:      31/01/2000
'Purpose:   Returns the windows directory
'Notes:     I know it does use the API, but I needed a quick fix
'           but there will be a faster API version soon.

Dim strTmpVal As String

strTmpVal = Environ$("WINDIR") 'Windows sets this so it IS correct

If Right(strTmpVal, 1) <> "\" Then
    strTmpVal = strTmpVal & "\"
End If

WindowsDirectory = strTmpVal

End Function
Public Function TempDirectory() As String

'Author:    Neil Ramsbottom
'Date:      31/01/2000
'Purpose:   Returns the temp directory
'Notes:     I know it does use the API, but I needed a quick fix
'           but there will be a faster API version soon.

Dim strTmpVal As String

strTmpVal = Environ$("TEMP") 'Windows sets this so it IS correct

If Right(strTmpVal, 1) <> "\" Then
    strTmpVal = strTmpVal & "\"
End If

TempDirectory = strTmpVal

End Function
Public Function GetAppPath() As String

'Author:    Neil Ramsbottom
'Date:      31/01/2000
'Purpose:   Returns App.Path with a black slash ALWAYS (if app.path was root,
'           it would return "C:", so a filename will not work i.e "C:data.dat").
'           Yeah, that will work in explorer, but not in VB5

If Right(App.Path, 1) <> "\" Then
    GetAppPath = App.Path & "\"
Else
    GetAppPath = App.Path
End If
    
End Function
Public Function LoadResStr(intResId As Integer) As String

'Author:    Neil Ramsbottom
'Date:      02/02/2000
'Purpose:   The LoadResString function with error checking

'Just cos its got error supression

On Error Resume Next

LoadResStr = LoadResString(intResId)


End Function
