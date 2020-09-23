Attribute VB_Name = "Module2"
Public Sub saveThis(mn As String)
MsgBox mn
Open App.Path & "\" & mn & ".txt" For Output As 1
For a = 0 To 19
    For b = 0 To 18
        Write #1, currentMap(b, a),
    Next b
    Write #1, currentMap(19, a)
Next a
Close
MsgBox "saved"
End Sub

Public Sub loadThis(mn As String)
Open App.Path & "\" & mn & ".txt" For Input As 1
For a = 0 To 19
    For b = 0 To 19
        Input #1, currentMap(b, a)
    Next b
Next a
Close
End Sub
