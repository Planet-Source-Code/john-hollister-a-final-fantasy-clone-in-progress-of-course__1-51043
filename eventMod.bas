Attribute VB_Name = "eventMod"
'*****************************************************************
'this mod opens and closes files that need to be rewritten during
'runtime when an event occurs.  the event
'*****************************************************************



Global tmpMap(-1 To 20, -1 To 20) As Integer

Public Sub perform(action As Integer)
    'move suzie and old man
    If action = 1 Then
        Call loadTmpMap("myhome.txt")
        tmpMap(3, 10) = 164
        Call saveThis("myhome.txt")
        Call loadTmpMap("cabinx1y7.txt")
        tmpMap(14, 15) = 170
        tmpMap(13, 16) = 18
        tmpMap(14, 16) = 20
        Call saveThis("cabinx1y7.txt")
        speech(1, 7, 13, 16, 0) = "Suzie:  With me around, no one can take his treasure!"
        progressSpeech(1, 7, 13, 16, 0) = -1
        Dim a As Integer
        For a = 0 To 20
            speech(1, 7, 14, 16, a) = speech(1, 7, 14, 15, a)
            progressSpeech(1, 7, 14, 16, a) = progressSpeech(1, 7, 14, 15, a)
            For b = 0 To 3
                affects(1, 7, 14, 16, (a * 4) + b) = affects(1, 7, 14, 15, (a * 4) + b)
            Next b
        Next a
            
    End If
    
    
    'lumber jack builds bridge
    If action = 2 Then
        Call loadTmpMap("x2y7.txt")
        tmpMap(8, 11) = 2
        tmpMap(9, 11) = 2
        tmpMap(10, 11) = 2
        Call saveThis("x2y7.txt")
    End If
    
    
    'get main characters name
    If action = 3 Then
        charNames(0) = InputBox("Enter your name: ")
        Call loadMapInfo
        talked = True
    End If
    
    
    'kill off myhome.txt
    If action = 4 Then
        talked = False
        Call loadTmpMap("myhome.txt")
        tmpMap(12, 16) = 164
        tmpMap(8, 9) = 164
        tmpMap(14, 5) = 164
        Call saveThis("myhome.txt")
        
        Call loadTmpMap("storex3y7.txt")
        tmpMap(18, 10) = 11
        Call saveThis("storex3y7.txt")
        Call loadTmpMap("storex11y4.txt")
        tmpMap(11, 14) = 11
        tmpMap(8, 13) = 11
        Call saveThis("storex11y4.txt")
        Call loadTmpMap("storex13y11.txt")
        tmpMap(15, 16) = 91
        tmpMap(15, 14) = 11
        Call saveThis("storex13y11.txt")
    
    End If
    
    'new member joins party
    If action = 5 Then
        charNames(1) = InputBox("Enter a name for new member.")
        party(1) = 2
        fightTeam(1) = 2
        Call setYourInitStats(2)
        Call checkLevel(2)
        talked = True
        Call loadTmpMap("castlef1.txt")
        tmpMap(8, 4) = 240
        Call saveThis("castlef1.txt")
    End If
    
    
    
    
End Sub

'********************************************************
'run this at a new game to undo all changes to files from
'a previous game
'********************************************************
Public Sub undoPerform(action As Integer)
    'move suzie and old man
    If action = 1 Then
        Call loadTmpMap("myhome.txt")
        tmpMap(3, 10) = 18
        Call saveThis("myhome.txt")
        Call loadTmpMap("cabinx1y7.txt")
        tmpMap(14, 15) = 20
        tmpMap(13, 16) = 164
        tmpMap(14, 16) = 170
        Call saveThis("cabinx1y7.txt")
    End If
    
    'lumberjack bridge
    If action = 2 Then
        Call loadTmpMap("x2y7.txt")
        tmpMap(8, 11) = 61
        tmpMap(9, 11) = 0
        tmpMap(10, 11) = 59
        Call saveThis("x2y7.txt")
    End If
    
    If action = 4 Then
        Call loadTmpMap("myhome.txt")
        tmpMap(12, 16) = 14
        tmpMap(8, 9) = 211
        tmpMap(14, 5) = 22
        Call saveThis("myhome.txt")
        Call loadTmpMap("storex3y7.txt")
        tmpMap(18, 10) = 213
        Call saveThis("storex3y7.txt")
        Call loadTmpMap("storex11y4.txt")
        tmpMap(11, 14) = 24
        tmpMap(8, 13) = 16
        Call saveThis("storex11y4.txt")
        Call loadTmpMap("storex13y11.txt")
        tmpMap(15, 16) = 28
        tmpMap(15, 14) = 16
        Call saveThis("storex13y11.txt")
    End If
    
    If action = 5 Then
        Call loadTmpMap("castlef1.txt")
        tmpMap(8, 4) = 229
        Call saveThis("castlef1.txt")
    End If
    
End Sub


