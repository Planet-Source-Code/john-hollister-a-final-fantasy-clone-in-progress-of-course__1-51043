Attribute VB_Name = "treasureMod"
Global treasure(0 To 9, 0 To 9, 0 To 19, 0 To 19) As Integer

'*****************************************
'*  This puts treasure in all the chests *
'*****************************************

Private Const ADDGIL10 = 1
Private Const GOTPOTION = 2
Private Const ADDGIL50 = 3
Private Const GOTDAGGER = 4
Private Const GOTPHOENIX = 5
Private Const GOTETHER = 6
Private Const GOTTENT = 7
Private Const GOTLEATHERARMOR = 8

Public Sub setTreasure()

    'each number in the array is the address in the game
    'i.e.  mapx, mapy, charx, chary
    'current map and then current location in map
    'just adding this will set the treasure automatically
    
    treasure(1, 7, 18, 1) = ADDGIL10
    treasure(1, 7, 18, 18) = GOTPOTION
    treasure(1, 7, 18, 17) = GOTPOTION
    treasure(1, 7, 16, 16) = ADDGIL50
    treasure(1, 7, 5, 5) = ADDGIL50
    treasure(1, 7, 6, 4) = GOTPOTION
    treasure(1, 7, 6, 3) = GOTPOTION
    treasure(1, 7, 4, 5) = GOTPHOENIX
    treasure(1, 7, 3, 5) = GOTETHER
    treasure(1, 7, 2, 5) = GOTTENT
    treasure(1, 7, 1, 5) = GOTPOTION
    treasure(1, 7, 17, 18) = GOTDAGGER
    treasure(2, 7, 4, 7) = GOTLEATHERARMOR
End Sub


Public Sub getTreasure(action As Integer)
    If action = ADDGIL10 Then
        gameScreen.dialogue.Caption = "Found 10 gil!"
        speech(1, 7, 14, 5, 0) = "Stinky Jill:  WAHHH! My money's gone again! Did you tell Mean Jean?!"
        gil = gil + 10
    End If
    
    If action = GOTPOTION Then
        gameScreen.dialogue.Caption = "Found potion!"
        For a = 0 To 300
            If itemInv(0, a) = "Potion" Then
                If itemInv(1, a) < 99 Then
                    itemInv(1, a) = itemInv(1, a) + 1
                End If
                a = 305
            End If
            If a <> 305 Then
                If itemInv(0, a) = "" Then
                    itemInv(0, a) = "Potion"
                    itemInv(1, a) = 1
                    a = 305
                End If
            End If
        Next a
    End If
    
    If action = ADDGIL50 Then
        gameScreen.dialogue.Caption = "Found 50 gil!"
        gil = gil + 50
    End If
    
    If action = GOTDAGGER Then
        gameScreen.dialogue.Caption = "Found a dagger!"
        For a = 0 To 50
            If weaponInv(NAMED, 1, a) = "" Then
                weaponInv(NAMED, 1, a) = "Dagger"
                weaponInv(POW, 1, a) = 4
                weaponInv(OFFDEF, 1, a) = OFFENSE
                a = 51
            End If
        Next a
    End If
    
    If action = GOTPHOENIX Then
        gameScreen.dialogue.Caption = "Found phoenix down!"
        For a = 0 To 300
            If itemInv(0, a) = "Phoenix" Then
                If itemInv(1, a) < 99 Then
                    itemInv(1, a) = itemInv(1, a) + 1
                End If
                a = 305
            End If
            If a <> 305 Then
                If itemInv(0, a) = "" Then
                    itemInv(0, a) = "Phoenix"
                    itemInv(1, a) = 1
                    a = 305
                End If
            End If
        Next a
    End If
    
    If action = GOTETHER Then
        gameScreen.dialogue.Caption = "Found ether!"
        For a = 0 To 300
            If itemInv(0, a) = "Ether" Then
                If itemInv(1, a) < 99 Then
                    itemInv(1, a) = itemInv(1, a) + 1
                End If
                a = 305
            End If
            If a <> 305 Then
                If itemInv(0, a) = "" Then
                    itemInv(0, a) = "Ether"
                    itemInv(1, a) = 1
                    a = 305
                End If
            End If
        Next a
    End If
        
    
    If action = GOTTENT Then
        gameScreen.dialogue.Caption = "Found a tent!"
        For a = 0 To 300
            If itemInv(0, a) = "Tent" Then
                If itemInv(1, a) < 99 Then
                    itemInv(1, a) = itemInv(1, a) + 1
                End If
                a = 305
            End If
            If a <> 305 Then
                If itemInv(0, a) = "" Then
                    itemInv(0, a) = "Tent"
                    itemInv(1, a) = 1
                    a = 305
                End If
            End If
        Next a
    End If
    
    
    If action = GOTLEATHERARMOR Then
        gameScreen.dialogue.Caption = "Found leather armor!"
        For a = 0 To 50
            If weaponInv(NAMED, 1, a) = "" Then
                weaponInv(NAMED, 1, a) = "Leth Armr"
                weaponInv(POW, 1, a) = 4
                weaponInv(OFFDEF, 1, a) = DEFENSE
                a = 51
            End If
        Next a

    End If
End Sub
