Attribute VB_Name = "MiscUtilities"
'Misc Utilities
'Jonathan Orgel, kalonymous, Inc.
Const VSS_ID = "$Header: /FastTextStream/Modules/MiscUtilities.bas 1     1/10/99 11:48a Joni $"
'$NoKeywords: $

Option Explicit

'Unload all forms in the current application
Public Sub UnloadAllForms()
    Dim Form As Form
    
    For Each Form In Forms
        Unload Form
        Set Form = Nothing 'This should not be necessary...
    Next Form
End Sub



Public Function FileExists(Filename As String) As Boolean
    'Fast check whether a file exists
    
    On Error GoTo FileDoesNotExist
    
    'Filelen will cause an error if the file does not exist
    FileLen (Filename)
    FileExists = True
    Exit Function
    
FileDoesNotExist:
    FileExists = False
End Function
