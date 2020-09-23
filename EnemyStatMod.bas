Attribute VB_Name = "EnemyStatMod"
Global enemyStats(0 To 1000, 0 To 14) As Integer
Global ename(0 To 1000) As String

Global Const FIRE = 0
Global Const ICE = 1
Global Const EARTH = 2
Global Const LIGHT = 3


Global Const CURE = 25




Global Const HP = 0
Global Const CURHP = 1
Global Const MP = 2
Global Const CURMP = 3
Global Const POWER = 4
Global Const DEFENSE = 5
Global Const MAGPOWER = 6
Global Const MAGDEFENSE = 7
Global Const STRENGTHS = 8
Global Const WEAKNESSES = 9
Global Const SPEED = 10
Global Const EXP = 12
Global Const GP = 11
Global Const LEVEL = 13

'***************************************************************
'* this sets all stats for enemies.. the rest is self-explanatory
'******************************************************************

Public Sub getStats(bad As Integer)
'aqua worm
    If bad = 0 Then
        
    End If
    

'armadillo
    If bad = 1 Then
        enemyStats(bad, HP) = 120
        enemyStats(bad, POWER) = 5
        enemyStats(bad, MP) = 15
        enemyStats(bad, GP) = 50
        enemyStats(bad, EXP) = 10
        enemyStats(bad, DEFENSE) = 50
        enemyStats(bad, WEAKNESSES) = FIRE
        enemyStats(bad, STRENGTHS) = EARTH
        enemyStats(bad, SPEED) = 5
        ename(bad) = "Armadillo"
    End If


'balloon
    If bad = 2 Then
        enemyStats(bad, POWER) = 5
        enemyStats(bad, HP) = 68
        enemyStats(bad, MP) = 20
        enemyStats(bad, GP) = 20
        enemyStats(bad, EXP) = 5
        enemyStats(bad, DEFENSE) = 100 'means weak!!!
        enemyStats(bad, WEAKNESSES) = LIGHT
        enemyStats(bad, STRENGTHS) = EARTH
        ename(bad) = "Balloon"
        enemyStats(bad, SPEED) = 8
    End If

'basilik
    If bad = 3 Then
        enemyStats(bad, HP) = 95
        enemyStats(bad, POWER) = 5
        enemyStats(bad, MP) = 5
        enemyStats(bad, GP) = 30
        enemyStats(bad, EXP) = 8
        enemyStats(bad, DEFENSE) = 100 'means weak!!
        ename(bad) = "Basilik"
        enemyStats(bad, SPEED) = 10
        enemyStats(bad, WEAKNESSES) = EARTH
        enemyStats(bad, STRENGTHS) = FIRE
    End If


'universal copying
        enemyStats(bad, CURHP) = enemyStats(bad, HP)
        enemyStats(bad, CURMP) = enemyStats(bad, MP)

End Sub

