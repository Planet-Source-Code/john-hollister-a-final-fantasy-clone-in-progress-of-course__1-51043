VERSION 5.00
Begin VB.Form Battle 
   Caption         =   "Battle!"
   ClientHeight    =   8265
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer updater 
      Interval        =   1
      Left            =   5280
      Top             =   6360
   End
   Begin VB.Timer endTime 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   5280
      Top             =   6960
   End
   Begin VB.Timer enemyAtk 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3840
      Top             =   6960
   End
   Begin VB.Timer atkMov 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4560
      Top             =   6960
   End
   Begin VB.Timer atbTimer 
      Interval        =   100
      Left            =   4560
      Top             =   6360
   End
   Begin VB.PictureBox statFrame 
      BackColor       =   &H80000007&
      Height          =   1575
      Left            =   240
      ScaleHeight     =   1515
      ScaleWidth      =   4335
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4680
      Width           =   4400
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000A&
         Height          =   975
         Left            =   240
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Special"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Magic"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Attack"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label member 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label member 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label member 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
      Begin VB.Shape atbBorder 
         BorderColor     =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   3300
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   975
      End
      Begin VB.Shape atb 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         Height          =   75
         Index           =   2
         Left            =   3360
         Top             =   1140
         Width           =   855
      End
      Begin VB.Shape atbBorder 
         BorderColor     =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   3300
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   975
      End
      Begin VB.Shape atb 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         Height          =   75
         Index           =   1
         Left            =   3360
         Top             =   660
         Width           =   855
      End
      Begin VB.Shape atbBorder 
         BorderColor     =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   3300
         Shape           =   4  'Rounded Rectangle
         Top             =   180
         Width           =   975
      End
      Begin VB.Shape atb 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         Height          =   75
         Index           =   0
         Left            =   3360
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label hpmp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   15
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label hpmp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   14
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label hpmp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   13
      Top             =   360
      Width           =   2535
   End
   Begin VB.Image magAtk 
      Height          =   615
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label hitDisp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label statList 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   435
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   4245
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Image team 
      Height          =   615
      Index           =   5
      Left            =   600
      Top             =   3000
      Width           =   615
   End
   Begin VB.Image team 
      Height          =   615
      Index           =   4
      Left            =   1800
      Top             =   2160
      Width           =   615
   End
   Begin VB.Image team 
      Height          =   615
      Index           =   3
      Left            =   600
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image team 
      Height          =   615
      Index           =   0
      Left            =   3840
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image team 
      Height          =   615
      Index           =   1
      Left            =   3840
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image team 
      Height          =   615
      Index           =   2
      Left            =   3840
      Top             =   3480
      Width           =   615
   End
   Begin VB.Image indicator 
      Height          =   240
      Left            =   3960
      Picture         =   "Battle.frx":0000
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4400
      Left            =   240
      Picture         =   "Battle.frx":0372
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4400
   End
End
Attribute VB_Name = "Battle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public inBattle, inMagic, inMenu, inWaiting, inItem, done As Boolean
Public currSlot, currMenu, currItem As Integer
Private takeTurn(0 To 5, 0 To 1) As Integer  ' 0 is for current, 1 is for max
Private dead(0 To 5) As Boolean
Private espeed(3 To 5) As Integer
Private eturn(3 To 5) As Integer
Public turn As Integer
Public menuSel As Integer
Public isAttack, isItem, isMagic As Boolean
Public magicType As String
Public itemType As String
Public totEXP, totGIL As Integer
Public weak, strong As Integer
Private tmpdead(0 To 2) As Boolean






Private Sub atbTimer_Timer()
    If iswaiting Then
        statList.Visible = False
    End If
    
    'take speed into account
    For a = 0 To 5
        If usedSlots(a) Then
            If a < 3 Then
                atb(a).Width = atb(a).Width + yourStats(fightTeam(a), SPEED)
            Else
                If usedSlots(a) And dead(a) = False Then
                    espeed(a) = espeed(a) + estats(a, etype(a), SPEED)
                    If espeed(a) >= 1000 Then
                        'attack player
                        Dim pick As Integer
                        pick = Int(Rnd * 3)
                        While usedSlots(pick) = False
                            pick = (pick + 1) Mod 3
                        Wend
                        Dim ehit As Integer
                        enemyAtk.Enabled = True
                        eturn(a) = a
                        ehit = Int(Rnd * estats(a, etype(a), POWER)) + estats(a, etype(a), POWER)
                        
                        hitDisp.Left = team(pick).Left + (team(pick).Width / 2)
                        hitDisp.Top = team(pick).Top - 200
                        hitDisp.Caption = ehit
                        hitDisp.Visible = True
                        yourStats(fightTeam(pick), CURHP) = Int(yourStats(fightTeam(pick), CURHP)) - Int(ehit)
                                                
                        'did attack kill you??
                        If yourStats(fightTeam(pick), CURHP) <= 0 Then
                            dead(pick) = True
                            
                            'change to dead picture
                            'team(pick).Visible = False
                            
                            
                            Dim battlelose As Boolean
                            battlelose = True
                            For b = 0 To 2
                                If dead(b) = False Then
                                    battlelose = False
                                    b = 6
                                End If
                            Next b
                            If battlelose Then
                                MsgBox "you lose!!!"
                                atbTimer.Enabled = False
                                endTime.Enabled = True
                            End If
                        End If
                        espeed(a) = 0
                    End If
                End If
            End If
            
            'add speed
            'temporary code
        End If
    Next a
    For a = 0 To 2
        If usedSlots(a) And atkMov.Enabled = False Then
            'update atb bar
            'temp code
            If atb(a).Width >= 855 Then
                turn = a
                statList.Visible = True
                statList.Caption = charNames(fightTeam(turn) - 1)
                If atb(a).Width = 855 Then
                    atb(a).Width = 855
                End If
                atb(a).BackColor = vbYellow
                member(a).ForeColor = vbYellow
                inMenu = True
                'for attack only
                menuItem(0).BackColor = vbRed
                atbTimer.Enabled = False
                a = 3
            End If
        End If
    Next a
        
End Sub

Private Sub atkMov_Timer()

    If team(turn).Left = 3840 Then
        team(turn).Left = 3600
        team(turn).Picture = people.guy((party(turn) * 8) - 1).Picture
    Else
        team(turn).Left = 3840
        team(turn).Picture = people.guy((party(turn) * 8) - 5).Picture
        atkMov.Enabled = False
        'animate attack movement
        hitDisp.Visible = False
        hitDisp.ForeColor = vbWhite
        
        
            isAttack = False
            isItem = False
            isMagic = False
            inMagic = False
            inItem = False
            currItem = 0
            currMenu = 0
            
        'only for if attacking enemy, not team
        If currSlot > 2 Then
            If estats(currSlot, etype(currSlot), CURHP) <= 0 Then
                dead(currSlot) = True
                'add effects if you want
                team(currSlot).Visible = False
                usedSlots(currSlot) = False
                While usedSlots(currSlot) = False
                    currSlot = (currSlot + 1) Mod 6
                Wend
                indicator.Left = team(currSlot).Left + (team(currSlot).Width / 2)
                indicator.Top = team(currSlot).Top - indicator.Height
                Dim battlewin As Boolean
                battlewin = True
                For a = 3 To 5
                    If dead(a) = False Then
                        battlewin = False
                        a = 6
                    End If
                Next a
                If battlewin Then
                    endTime.Enabled = True
                End If
            End If
        Else
        'attacking yourself!!
            If yourStats(fightTeam(currSlot), CURHP) <= 0 Then
                dead(currSlot) = True
                yourStats(fightTeam(currSlot), CURHP) = 0
                'add effects if you want
                
                'add dead picture here!
                'team(currSlot).Visible = False
                usedSlots(currSlot) = False
                While usedSlots(currSlot) = False
                    currSlot = (currSlot + 1) Mod 6
                Wend
                indicator.Left = team(currSlot).Left + (team(currSlot).Width / 2)
                indicator.Top = team(currSlot).Top - indicator.Height
                Dim battlelose As Boolean
                battlelose = True
                For a = 0 To 2
                    If dead(a) = False Then
                        battlelose = False
                        a = 6
                    End If
                Next a
                If battlelose Then
                    MsgBox "you lose!!!"
                    endTime.Enabled = True
                End If
            End If
        End If
    End If
    magAtk.Visible = False
End Sub

Private Sub enemyAtk_Timer()
    Dim ego As Integer
    ego = -1
    For a = 3 To 5
        If eturn(a) <> 0 Then
            ego = a
            a = 6
        End If
    Next a
    
    If ego = -1 Then
        enemyAtk.Enabled = False
    Else
        If team(3).Left = 600 And ego = 3 Then
            team(ego).Left = 840
            hitDisp.Visible = True
        Else
            If ego = 3 Then
                team(ego).Left = 600
                eturn(ego) = 0
                hitDisp.Visible = False
            End If
        End If
        
        If team(4).Left = 1800 And ego = 4 Then
            team(ego).Left = 2040
            hitDisp.Visible = True
        Else
            If ego = 4 Then
                team(ego).Left = 1800
                hitDisp.Visible = False
                eturn(ego) = 0
            End If
        End If
        
        If team(5).Left = 600 And ego = 5 Then
            team(ego).Left = 840
            hitDisp.Visible = True
        Else
            If ego = 5 Then
                team(ego).Left = 600
                eturn(ego) = 0
                hitDisp.Visible = False
            End If
        End If
    End If


End Sub



Private Sub statFrame_KeyDown(KeyCode As Integer, Shift As Integer)
    'selecting enemy
    
    
    
    
    
    Dim tmpf As Integer
    Dim b As Integer
        
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyDown Then
        tmpf = 1
    End If
        If KeyCode = vbKeyUp Or KeyCode = vbKeyRight Then
        tmpf = -1
    End If


    
    
    'battle stuff
    If inBattle Then
        If tmpf = -1 And currSlot = 0 Then
            currSlot = 5
        Else
            currSlot = (currSlot + tmpf) Mod 6
        End If
        While usedSlots(currSlot) = False
            If tmpf = -1 And currSlot = 0 Then
                currSlot = 5
            Else
                currSlot = (currSlot + tmpf) Mod 6
            End If
        Wend
        indicator.Left = team(currSlot).Left + (team(currSlot).Width / 2)
        indicator.Top = team(currSlot).Top - indicator.Height
        
        'display stats in label!!!
        If currSlot < 3 Then
            statList.Caption = charNames(fightTeam(currSlot) - 1)
        Else
            statList.Caption = ename(etype(currSlot))
        End If

    End If
    

    
    'menu screen controls
    If inMenu Then
        menuSel = 0
        menuItem(currMenu).BackColor = &H8000000D
        If (KeyCode = vbKeyUp Or KeyCode = vbKeyRight) And currMenu = 0 Then
            currMenu = 3
        Else
            currMenu = Abs(currMenu + tmpf) Mod 4
        End If
        menuItem(currMenu).BackColor = vbRed
    End If
    
    
    
    
    'selecting an item
    If inItem Then
        menuItem(currItem).BackColor = &H8000000D
        If menuSel + tmpf < 301 And tmpf > 0 Then
            If itemInv(0, menuSel + tmpf) <> "" Then
                menuSel = menuSel + tmpf
                If menuSel > 3 And currItem = 3 And tmpf = 1 Then
                    b = 3
                    For a = menuSel To menuSel - 3 Step -1
                        menuItem(b).Caption = itemInv(0, a) & "... " & itemInv(1, a)
                        b = b - 1
                    Next a
                End If
                If currItem < 3 And tmpf = 1 Then
                    currItem = currItem + 1
                End If
            End If
        Else
            If menuSel > 0 And tmpf < 0 Then
                If itemInv(0, menuSel + tmpf) <> "" Then
                    If menuSel > 0 And currItem = 0 Then
                        b = 0
                        For a = menuSel - 1 To menuSel + 2
                            menuItem(b).Caption = itemInv(0, a) & "... " & itemInv(1, a)
                            b = b + 1
                        Next a
                    End If
                    menuSel = menuSel + tmpf
                    If currItem > 0 Then
                        currItem = currItem - 1
                    End If
                End If
            End If
        End If
        menuItem(currItem).BackColor = vbRed
    End If
    
    
     If inMagic Then
        menuItem(currItem).BackColor = &H8000000D
        If menuSel + tmpf < 301 And tmpf > 0 Then
                                'CHANGED FROM CURRITEM - IN OTHER ONE TOO ^
            If magicInv(0, fightTeam(turn), menuSel + tmpf) <> "" Then
                menuSel = menuSel + tmpf
                If menuSel > 3 And currItem = 3 And tmpf = 1 Then
                    b = 3
                    For a = menuSel To menuSel - 3 Step -1
                        menuItem(b).Caption = magicInv(0, fightTeam(turn), a) & "... " & magicInv(1, fightTeam(turn), a)
                        b = b - 1
                    Next a
                End If
                If currItem < 3 And tmpf = 1 Then
                    currItem = currItem + 1
                End If
            End If
        Else
            If menuSel > 0 And tmpf < 0 Then
                If magicInv(0, fightTeam(turn), menuSel + tmpf) <> "" Then
                    If menuSel > 0 And currItem = 0 Then
                        b = 0
                        For a = menuSel - 1 To menuSel + 2
                            menuItem(b).Caption = magicInv(0, fightTeam(turn), a) & "... " & magicInv(1, fightTeam(turn), a)
                            b = b + 1
                        Next a
                    End If
                    menuSel = menuSel + tmpf
                    If currItem > 0 Then
                        currItem = currItem - 1
                    End If
                End If
            End If
        End If
        menuItem(currItem).BackColor = vbRed
    End If
   
    
    
    
    If KeyCode = vbKeyTab Then
        indicator.Visible = False
        If inBattle Then
            statList.Visible = False
        End If
        If inBattle And (isMagic Or isItem) Then
            menuItem(currItem).BackColor = vbRed
            inBattle = False
            If isMagic Then
                inMagic = True
                isMagic = False
            End If
            If isItem Then
                For a = 0 To 2
                    If fightTeam(a) <> 0 And itemType = "Phoenix" Then
                        usedSlots(a) = tmpdead(a)
                    End If
                Next a
                inItem = True
                isItem = False
            End If
        Else
            If inItem Or inMagic Or (inBattle And isAttack) Then
                inItem = False
                For a = 0 To 3
                    menuItem(a).BackColor = &H8000000D
                Next a
                menuItem(0).Caption = "Attack"
                menuItem(1).Caption = "Magic"
                menuItem(2).Caption = "Item"
                menuItem(3).Caption = "Special"
                
                menuItem(currMenu).BackColor = vbRed
                inBattle = False
                isAttack = False
                inMenu = True
                inMagic = False
                inItem = False
                currItem = 0
            End If
        End If
    End If
    
    
    
    
    
    
    
    
    
    'might have to add to this
    If KeyCode = vbKeySpace Then

    
        If inBattle Then
            indicator.Visible = False
            inBattle = False
            inWaiting = True
            atb(turn).BackColor = vbRed
            atb(turn).Width = 10
            atbTimer.Enabled = True
            atkMov.Enabled = True
            
            Dim rhit As Double
            Dim totHit As Integer
            
            'perform attacks
            
'*******************************************************
'*          NORMAL ATTACK NORMAL ATTACK NORMAL         *
'*******************************************************
            If isAttack Then
                'REGULAR ATTACK
                Randomize
                rhit = (Rnd * yourStats(fightTeam(turn), LEVEL) - (yourStats(fightTeam(turn), LEVEL) / 2))
                totHit = ((yourStats(fightTeam(turn), POWER) + offeq(POW, fightTeam(turn))) * (yourStats(fightTeam(turn), LEVEL))) + rhit
                'calculate defense
                If currSlot > 2 Then
                    totHit = totHit / (100 / estats(currSlot, etype(currSlot), DEFENSE))
                End If
                'tothit is the final blow total
                menuItem(currMenu).BackColor = &H8000000D
                member(turn).ForeColor = vbWhite
                hitDisp.Left = team(currSlot).Left + (team(currSlot).Width / 2)
                hitDisp.Top = team(currSlot).Top - 300
                hitDisp.Caption = totHit
                hitDisp.Visible = True
                If currSlot > 2 Then
                    estats(currSlot, etype(currSlot), CURHP) = estats(currSlot, etype(currSlot), CURHP) - totHit
                Else
                    yourStats(fightTeam(currSlot), CURHP) = yourStats(fightTeam(currSlot), CURHP) - totHit
                End If
            End If
            
            
            
            
'*******************************************************
'*          MAGIC ATTACK MAGIC MAGIC ATTAck            *
'*******************************************************
            
            If isMagic Then
                'MAGIC ATTACK
                magAtk.Left = team(currSlot).Left
                magAtk.Top = team(currSlot).Top
                If magicType = "Cure" Then
                    weak = CURE
                End If
                If magicType = "Fire" Then
                    weak = FIRE
                    magAtk.Picture = people.magicAtk(0).Picture
                    magAtk.Visible = True
                End If
                If magicType = "Ice" Then
                    weak = ICE
                    magAtk.Picture = people.magicAtk(1).Picture
                    magAtk.Visible = True
                End If
                If magicType = "Light" Then
                    weak = LIGHT
                    magAtk.Picture = people.magicAtk(2).Picture
                    magAtk.Visible = True
                End If
                If magicType = "Earth" Then
                    weak = EARTH
                    magAtk.Picture = people.magicAtk(3).Picture
                    magAtk.Visible = True
                End If
                Randomize
                rhit = 100 + ((Rnd * 30) - 15)
                yourStats(fightTeam(turn), CURMP) = yourStats(fightTeam(turn), CURMP) - magicInv(1, fightTeam(turn), menuSel)
                totHit = Int(((yourStats(fightTeam(turn), MAGPOWER) + 50) * (rhit / 100)))
                
                
                'calculate weakness/strengths

                
                If currSlot > 2 Then
                    If estats(currSlot, etype(currSlot), WEAKNESSES) = weak Then
                        totHit = totHit * 2
                    End If
                    If estats(currSlot, etype(currSlot), STRENGTHS) = weak Then
                        totHit = totHit / 2
                    End If
                End If
                
                'tothit is the final blow total
                menuItem(currMenu).BackColor = &H8000000D
                member(turn).ForeColor = vbWhite

                If currSlot > 2 And weak <> CURE Then
                    estats(currSlot, etype(currSlot), CURHP) = estats(currSlot, etype(currSlot), CURHP) - totHit
                End If
                If currSlot < 3 And weak <> CURE Then
                    yourStats(fightTeam(currSlot), CURHP) = yourStats(fightTeam(currSlot), CURHP) - totHit
                End If
                
                'cure spells
                If currSlot > 2 And weak = CURE Then
                    estats(currSlot, etype(currSlot), CURHP) = estats(currSlot, etype(currSlot), CURHP) - totHit
                End If
                If currSlot < 3 And weak = CURE Then
                    hitDisp.ForeColor = vbGreen
                    If yourStats(fightTeam(currSlot), CURHP) < yourStats(fightTeam(currSlot), HP) Then
                        yourStats(fightTeam(turn), CURMP) = yourStats(fightTeam(turn), CURMP)
                        totHit = (yourStats(fightTeam(turn), MAGPOWER) * (yourStats(fightTeam(turn), LEVEL) / 2))
                        yourStats(fightTeam(currSlot), CURHP) = yourStats(fightTeam(currSlot), CURHP) + totHit
                        If yourStats(fightTeam(currSlot), CURHP) > yourStats(fightTeam(currSlot), HP) Then
                            totHit = totHit - (yourStats(fightTeam(currSlot), CURHP) - yourStats(fightTeam(currSlot), HP))
                            yourStats(fightTeam(currSlot), CURHP) = yourStats(fightTeam(currSlot), HP)
                        End If
                    End If
                End If
            
            
            
            
                hitDisp.Left = team(currSlot).Left + (team(currSlot).Width / 2)
                hitDisp.Top = team(currSlot).Top - 300
                hitDisp.Caption = totHit
                hitDisp.Visible = True
            
            
            
            End If
            
            
            
            
            
'*******************************************************
'*          ITEM ITEM APPLY AN ITEM USE ITEM           *
'*******************************************************
            
            If isItem Then
            
                'USE rhit
                If itemType = "Ether" Then
                    hitDisp.ForeColor = vbYellow
                    rhit = 0
                    mhit = 50
                End If
                
                If itemType = "Potion" Then
                    hitDisp.ForeColor = vbGreen
                    rhit = 50
                End If
                
                If itemType = "Phoenix" Then
                    hitDisp.ForeColor = vbGreen
                    For a = 0 To 2
                        If fightTeam(a) <> 0 Then
                            usedSlots(a) = tmpdead(a)
                        End If
                    Next a
                    If dead(currSlot) Then
                        rhit = Int(yourStats(fightTeam(currSlot), HP) / 10)
                        dead(currSlot) = False
                        usedSlots(currSlot) = True
                    Else
                        rhit = 0
                    End If
                End If

                'remove used item from inventory
                For a = 0 To 299
                    If itemInv(0, a) = itemType Then
                        itemInv(1, a) = itemInv(1, a) - 1
                        If itemInv(1, a) = 0 Then
                            For b = menuSel To 299
                                itemInv(0, b) = itemInv(0, b + 1)
                                itemInv(1, b) = itemInv(1, b + 1)
                            Next b
                        End If
                        a = 300
                    End If
                Next a
                'add and check hp limits
                totHit = rhit
                If totHit = 0 Then
                    totHit = mhit
                End If

                'POTIONS HP
                If currSlot < 3 And rhit <> 0 And dead(currSlot) = False Then
                    yourStats(fightTeam(currSlot), CURHP) = yourStats(fightTeam(currSlot), CURHP) + totHit
                    If yourStats(fightTeam(currSlot), CURHP) > yourStats(fightTeam(currSlot), HP) Then
                        totHit = totHit - (yourStats(fightTeam(currSlot), CURHP) - yourStats(fightTeam(currSlot), HP))
                        yourStats(fightTeam(currSlot), CURHP) = yourStats(fightTeam(currSlot), HP)
                    End If
                Else
                    If currSlot > 2 Then
                        estats(currSlot, etype(currSlot), CURHP) = estats(currSlot, etype(currSlot), CURHP) + totHit
                        If estats(currSlot, etype(currSlot), CURHP) > estats(currSlot, etype(currSlot), HP) Then
                            totHit = totHit - ((estats(currSlot, etype(currSlot), CURHP) - estats(currSlot, etype(currSlot), HP)))
                            estats(currSlot, etype(currSlot), CURHP) = estats(currSlot, etype(currSlot), HP)
                        End If
                    End If
                End If
                
                
                'ETHERS MP
                If currSlot < 3 And mhit <> 0 And dead(currSlot) = False Then
                    yourStats(fightTeam(currSlot), CURMP) = yourStats(fightTeam(currSlot), CURMP) + totHit
                    If yourStats(fightTeam(currSlot), CURMP) > yourStats(fightTeam(currSlot), MP) Then
                        totHit = totHit - (yourStats(fightTeam(currSlot), CURMP) - yourStats(fightTeam(currSlot), MP))
                        yourStats(fightTeam(currSlot), CURMP) = yourStats(fightTeam(currSlot), MP)
                    End If
                Else
                    If currSlot > 2 Then
                        estats(currSlot, etype(currSlot), CURMP) = estats(currSlot, etype(currSlot), CURMP) + totHit
                        If estats(currSlot, etype(currSlot), CURMP) > estats(currSlot, etype(currSlot), MP) Then
                            totHit = totHit - ((estats(currSlot, etype(currSlot), CURMP) - estats(currSlot, etype(currSlot), MP)))
                            estats(currSlot, etype(currSlot), CURMP) = estats(currSlot, etype(currSlot), MP)
                        End If
                    End If
                End If
                
                'tothit is the final blow total
                menuItem(currMenu).BackColor = &H8000000D
                member(turn).ForeColor = vbWhite
                hitDisp.Left = team(currSlot).Left + (team(currSlot).Width / 2)
                hitDisp.Top = team(currSlot).Top - 300
                hitDisp.Caption = totHit
                hitDisp.Visible = True
                mhit = 0
                rhit = 0
            End If
            

            statList.Visible = False
            
            For a = 0 To 3
                menuItem(a).BackColor = &H8000000D
            Next a
            menuItem(0).Caption = "Attack"
            menuItem(1).Caption = "Magic"
            menuItem(2).Caption = "Item"
            menuItem(3).Caption = "Special"

'*****************************************************************
           
            
            
        Else
            If inMenu Then
                While usedSlots(currSlot) = False
                    currSlot = (currSlot + 1) Mod 6
                Wend
                ' selects normal attack
                If currMenu = 0 Then
                    indicator.Left = team(currSlot).Left + (team(currSlot).Width / 2)
                    indicator.Top = team(currSlot).Top - indicator.Height
                    indicator.Visible = True
                    inBattle = True
                    'display stats in label!!!
                    statList.Caption = charNames(currSlot)
                    inMenu = False
                    isAttack = True
                    statList.Visible = True
                End If
                
                'magic!!!
                If currMenu = 1 Then
                    For a = 0 To 3
                        menuItem(a).BackColor = &H8000000D
                        menuItem(a).Caption = magicInv(0, fightTeam(turn), a) & "... " & magicInv(1, fightTeam(turn), a)
                    Next a
                    menuItem(0).BackColor = vbRed
                    inMagic = True
                    inMenu = False
                End If
                
                'items!!!
                If currMenu = 2 Then
                    For a = 0 To 3
                        menuItem(a).BackColor = &H8000000D
                        menuItem(a).Caption = itemInv(0, a) & "... " & itemInv(1, a)
                    Next a
                    menuItem(0).BackColor = vbRed
                    inItem = True
                    inMenu = False
                End If
                
                
                
                
                
                'for special attacks
                
                
                
                
                
                
            Else
                If inMagic And magicInv(0, fightTeam(turn), menuSel) <> "" And magicInv(1, fightTeam(turn), menuSel) > yourStats(fightTeam(turn), CURMP) Then
                    statList.Caption = "NOT ENOUGH MP!"
                    statList.Visible = True
                End If
                If inMagic And magicInv(0, fightTeam(turn), menuSel) <> "" And magicInv(1, fightTeam(turn), menuSel) <= yourStats(fightTeam(turn), CURMP) Then
                    statList.Caption = ""
                    statList.Visible = True
                    magicType = magicInv(0, fightTeam(turn), menuSel)
                    
                    inBattle = True
                    isMagic = True
                    indicator.Left = team(currSlot).Left + (team(currSlot).Width / 2)
                    indicator.Top = team(currSlot).Top - indicator.Height
                    indicator.Visible = True
                    inMagic = False
                End If
                If inItem And itemInv(0, menuSel) <> "" Then
                    statList.Caption = ""
                    statList.Visible = True
                    itemType = itemInv(0, menuSel)
                    
                    If itemType = "Phoenix" Then
                        For a = 0 To 2
                            If fightTeam(a) <> 0 Then
                                tmpdead(a) = usedSlots(a)
                                usedSlots(a) = True
                            End If
                        Next a
                    End If
                    
                    inBattle = True
                    isItem = True
                    indicator.Left = team(currSlot).Left + (team(currSlot).Width / 2)
                    indicator.Top = team(currSlot).Top - indicator.Height
                    indicator.Visible = True
                    inItem = False
                End If
            End If
        End If
    End If
    
    
    If KeyCode = vbKeyEscape Then
    totGIL = 0
    totEXP = 0
        gameScreen.Enabled = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Call makeBattle
inBattle = False
inMenu = False
inWaiting = True
currSlot = 0
menuSel = 0
currMenu = 0
totEXP = 0
totGIL = 0
done = False



For a = 0 To 2
    If usedSlots(a) = False Then
        atbBorder(a).Visible = False
        hpmp(a).Visible = False
        atb(a).Visible = False
    End If
    atb(a).Width = 15
    espeed(a + 3) = 10
    If usedSlots(a + 3) Then
        totEXP = totEXP + estats(a + 3, etype(a + 3), EXP)
        totGIL = totGIL + estats(a + 3, etype(a + 3), GP)
    End If
Next a


For a = 0 To 5
    If usedSlots(a) = True Then
        dead(a) = False
    Else
        dead(a) = True
    End If
Next a

End Sub

Private Sub endTime_Timer()
If done = False Then
    inMenu = False
    inBattle = False
    inMagic = False
    inItem = False
    

    
    atbTimer.Enabled = False
    atkMov.Enabled = False
    enemyAtk.Enabled = False
    statList.Visible = True
    statList.Caption = "Got : " & totEXP & " EXP and " & totGIL & " GP"
    
    gil = gil + totGIL
    Dim a As Integer
    
    For a = 0 To 2
        If dead(a) = False Then
            yourStats(fightTeam(a), EXP) = yourStats(fightTeam(a), EXP) + totEXP
            Call checkLevel(fightTeam(a))
        Else
            yourStats(fightTeam(a), CURHP) = 1
        End If
    Next
    
    
    totGIL = 0
    totEXP = 0
    done = True
Else
    
    totGIL = 0
    totEXP = 0
    gameScreen.Enabled = True
    Unload Me
End If
End Sub

Private Sub updater_Timer()

    For a = 0 To 2
        If fightTeam(a) <> 0 Then
            hpmp(a).Caption = charNames(fightTeam(a) - 1) & "   " & yourStats(fightTeam(a), CURHP) & "/" & yourStats(fightTeam(a), HP) & "   " & yourStats(fightTeam(a), CURMP) & "/" & yourStats(fightTeam(a), MP)
        End If
    Next a

End Sub
