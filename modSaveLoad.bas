Attribute VB_Name = "modSaveLoad"
Public Sub saveTreasureState()
Open App.Path & "\screens\" & mapFile For Output As 1
For a = 0 To 19
    For b = 0 To 18
        Write #1, currentMap(b, a),
    Next b
    Write #1, currentMap(19, a)
Next a
Close
End Sub
Public Sub saveThis(mname As String)
Open App.Path & "\screens\" & mname For Output As 1
For a = 0 To 19
    For b = 0 To 18
        Write #1, tmpMap(b, a),
    Next b
    Write #1, tmpMap(19, a)
Next a
Close
End Sub

Public Sub loadMap(mname As String)
mapFile = mname
If mname <> "" Then
    Open App.Path & "\screens\" & mname For Input As 1
    
    For a = 0 To 19
        For b = 0 To 19
            Input #1, currentMap(b, a)
        Next b
    Next a
    Close
End If

End Sub


Public Sub loadTmpMap(mname As String)
mapFile = mname
If mname <> "" Then
    Open App.Path & "\screens\" & mname For Input As 1
    
    For a = 0 To 19
        For b = 0 To 19
            Input #1, tmpMap(b, a)
        Next b
    Next a
    Close
End If
End Sub


Public Sub resetTreasure()
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer

For a = 0 To 9
    For b = 0 To 9
        For c = 0 To 19
            For d = 0 To 19
                If mapHotSpot(a, b, c, d) <> "" Then
                    loadMap (mapHotSpot(a, b, c, d))
                    For e = 0 To 19
                        For f = 0 To 19
                            If currentMap(f, e) = 115 Then
                                currentMap(f, e) = 114
                                Call saveTreasureState
                            End If
                        Next f
                    Next e
                End If
                If hotDoor(a, b, c, d) <> "" Then
                    loadMap (hotDoor(a, b, c, d))
                    For e = 0 To 19
                        For f = 0 To 19
                            If currentMap(f, e) = 115 Then
                                currentMap(f, e) = 114
                                Call saveTreasureState
                            End If
                        Next f
                    Next e
                End If
            Next d
        Next c
    Next b
Next a


End Sub
