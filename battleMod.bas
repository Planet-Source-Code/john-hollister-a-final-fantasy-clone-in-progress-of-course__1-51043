Attribute VB_Name = "battleMod"
Global magicInv(0 To 1, 0 To 10, 0 To 20) As Variant
'0 to 1 is for name and mp usage
'0 to 10 is for owner

Global weaponInv(0 To 2, 0 To 10, 0 To 50) As Variant
'0 to 2 is for name and power and offense/defense





Global itemInv(0 To 1, 0 To 300) As Variant
'0 to 1 for name and amount
'0 to 300 would store names and amounts then

Private leveling(0 To 99) As Long


Global party(0 To 10) As Integer

Global fightTeam(0 To 2) As Integer

Global enemyTeam(0 To 2) As String

Global usedSlots(0 To 5) As Boolean

Global estats(0 To 5, 0 To 1000, 0 To 13) As Long

Global etype(3 To 5) As Integer

Global yourStats(0 To 10, 0 To 13) As Long

Global offeq(0 To 1, 0 To 10) As Variant
Global defeq(0 To 1, 0 To 10) As Variant
'0 is for named/ 1 is for power
Global Const CURWEAPON = 0
Global Const CURSHIELD = 1
Global Const POW = 1
Global Const NAMED = 0
Global Const OFFENSE = 2
Global Const OFFDEF = 2

'***********************************************************
'you can set up the initial inventory and player stats here
'***********************************************************
Public Sub setYourInitStats(mem As Integer)
gil = 1000
If mem = 1 Then
    
    itemInv(0, 0) = "Potion"
    itemInv(1, 0) = 2
    itemInv(0, 1) = "Ether"
    itemInv(1, 1) = 2
    
    magicInv(0, mem, 0) = "Fire"
    magicInv(1, mem, 0) = 4
    magicInv(0, mem, 1) = "Ice"
    magicInv(1, mem, 1) = 4
    magicInv(0, mem, 2) = "Light"
    magicInv(1, mem, 2) = 4
    magicInv(0, mem, 3) = "Earth"
    magicInv(1, mem, 3) = 4
    
    yourStats(mem, LEVEL) = 0
    yourStats(mem, EXP) = 200
    yourStats(mem, HP) = 100
    yourStats(mem, MP) = 30
    yourStats(mem, POWER) = 5
    yourStats(mem, MAGPOWER) = 20
    yourStats(mem, DEFENSE) = 18
    yourStats(mem, SPEED) = 15
    
    offeq(POW, mem) = 2
    offeq(NAMED, mem) = "Knife"
    
    defeq(POW, mem) = 2
    defeq(NAMED, mem) = "Leather"
    



End If
If mem = 2 Then

    magicInv(0, mem, 0) = "Cure"
    magicInv(1, mem, 0) = 6
    magicInv(0, mem, 1) = "Haste"
    magicInv(1, mem, 1) = 8
    
    yourStats(mem, LEVEL) = 0

    yourStats(mem, EXP) = 400
    yourStats(mem, HP) = 40
    yourStats(mem, MP) = 32
    yourStats(mem, POWER) = 20
    yourStats(mem, MAGPOWER) = 15
    yourStats(mem, DEFENSE) = 15
    yourStats(mem, SPEED) = 17
    
    
    offeq(POW, mem) = 6
    offeq(NAMED, mem) = "B. Sword"
    
    defeq(POW, mem) = 4
    defeq(NAMED, mem) = "Met. Vest"
    
    
    
End If
    
    yourStats(mem, CURHP) = yourStats(mem, HP)
    yourStats(mem, CURMP) = yourStats(mem, MP)




End Sub



'*********************************************************
'*  randomly choose a pre-set enemy party for battle
'*********************************************************

'have yet to make a way for enemies to occur distinctively to their location on map

Sub makeBattle()
Randomize
Dim tmpe As Integer
Dim a As Integer

For a = 0 To 5
    usedSlots(a) = False
Next a

'load enemy parties
tmpe = Int(Rnd * 4)
'multiply tmpe by offset to obtain parties for diff environments
If tmpe = 0 Then
    For a = 3 To 4
        Battle.team(a).Picture = people.baddie(3).Picture
        etype(a) = 3
        usedSlots(a) = True
        usedSlots(5) = False
        For b = 0 To 13
            estats(a, 3, b) = enemyStats(3, b)
        Next b
    Next a
End If
If tmpe = 1 Then
    For a = 4 To 5
        Battle.team(a).Picture = people.baddie(2).Picture
        usedSlots(a) = True
        etype(a) = 2
        For b = 0 To 13
            estats(a, 2, b) = enemyStats(2, b)
        Next b
    Next a
End If

If tmpe = 2 Then
    For a = 3 To 5
        Battle.team(a).Picture = people.baddie(1).Picture
        usedSlots(a) = True
        etype(a) = 1
        For b = 0 To 13
            estats(a, 1, b) = enemyStats(1, b)
        Next b
    Next a
End If
If tmpe = 3 Then
    For a = 4 To 5
        Battle.team(a).Picture = people.baddie(2).Picture
        usedSlots(a) = True
        etype(a) = 2
        For b = 0 To 13
            estats(a, 2, b) = enemyStats(2, b)
        Next b
    Next a
End If

For a = 0 To 2
    If party(a) <> 0 Then
        'load active party
        
        Battle.team(a).Picture = people.guy((party(a) * 8) - 5).Picture
        usedSlots(a) = True
        Battle.member(a).Caption = charNames(a)
    End If
Next a
End Sub


'*************************************************************
'* this method tests to see if a character has leveled up after
'* a battle, then the stats are updated if yes
'**************************************************************


Public Sub checkLevel(memb As Integer)
For a = 0 To 99
    leveling(a) = 0
Next a
    leveling(0) = 10
For a = 1 To 80
    leveling(a) = leveling(a - 1) + (100 * (1.5 * (2.71828 ^ (0.1 * (a - 15)))))
Next a
For a = 80 To 98
    leveling(a + 1) = (1.05) * leveling(a)
Next a
Dim tl As Integer
    For a = 0 To 99
    tl = yourStats(memb, LEVEL)
        If tl < 99 Then
            If yourStats(memb, EXP) > leveling(tl) Then
                yourStats(memb, LEVEL) = yourStats(memb, LEVEL) + 1
                If tl >= 5 And tl < 10 Then
                    yourStats(memb, HP) = yourStats(memb, HP) + 40
                    yourStats(memb, MP) = yourStats(memb, MP) + 4
                End If
                If tl >= 10 And tl < 30 Then
                    yourStats(memb, HP) = yourStats(memb, HP) + 60
                    yourStats(memb, MP) = yourStats(memb, MP) + 6
                End If
                If tl >= 30 And tl < 80 Then
                    yourStats(memb, HP) = yourStats(memb, HP) + 150
                    yourStats(memb, MP) = yourStats(memb, MP) + 15
                End If
            End If
        End If
    Next a

End Sub







