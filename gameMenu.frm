VERSION 5.00
Begin VB.Form gameMenu 
   Caption         =   "Menu"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4560
      Top             =   6240
   End
   Begin VB.PictureBox RPG 
      BackColor       =   &H00800000&
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   5475
      Left            =   240
      ScaleHeight     =   5415
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   240
      Width           =   4400
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000009&
         Height          =   5300
         Left            =   2800
         Top             =   50
         Width           =   1450
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000009&
         Height          =   735
         Left            =   45
         Top             =   4605
         Width           =   2775
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         Height          =   5295
         Left            =   45
         Top             =   45
         Width           =   2775
      End
      Begin VB.Label lvl 
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl"
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
         Index           =   2
         Left            =   1440
         TabIndex        =   16
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lvl 
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl"
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
         Index           =   1
         Left            =   1440
         TabIndex        =   15
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lvl 
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl"
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
         Index           =   0
         Left            =   1440
         TabIndex        =   14
         Top             =   120
         Width           =   855
      End
      Begin VB.Label totalgil 
         BackStyle       =   0  'Transparent
         Caption         =   "Gil"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label exper 
         BackStyle       =   0  'Transparent
         Caption         =   "exp"
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
         Index           =   2
         Left            =   1080
         TabIndex        =   12
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label myname 
         BackStyle       =   0  'Transparent
         Caption         =   "name"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label mpoints 
         BackStyle       =   0  'Transparent
         Caption         =   "mp"
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
         Index           =   2
         Left            =   1080
         TabIndex        =   10
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label hpoints 
         BackStyle       =   0  'Transparent
         Caption         =   "hp"
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
         Index           =   2
         Left            =   1080
         TabIndex        =   9
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Image team 
         Height          =   855
         Index           =   2
         Left            =   120
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label exper 
         BackStyle       =   0  'Transparent
         Caption         =   "exp"
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
         Index           =   1
         Left            =   1080
         TabIndex        =   8
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label myname 
         BackStyle       =   0  'Transparent
         Caption         =   "name"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label mpoints 
         BackStyle       =   0  'Transparent
         Caption         =   "mp"
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
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label hpoints 
         BackStyle       =   0  'Transparent
         Caption         =   "hp"
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
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Image team 
         Height          =   855
         Index           =   1
         Left            =   120
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label exper 
         BackStyle       =   0  'Transparent
         Caption         =   "exp"
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
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label myname 
         BackStyle       =   0  'Transparent
         Caption         =   "name"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label mpoints 
         BackStyle       =   0  'Transparent
         Caption         =   "MP: 999/999"
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
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label hpoints 
         BackStyle       =   0  'Transparent
         Caption         =   "hp"
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
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.Image team 
         Height          =   855
         Index           =   0
         Left            =   120
         Stretch         =   -1  'True
         Top             =   480
         Width           =   855
      End
      Begin VB.Label selects 
         BackColor       =   &H8000000D&
         Height          =   1455
         Index           =   0
         Left            =   50
         TabIndex        =   22
         Top             =   120
         Width           =   2780
      End
      Begin VB.Label selects 
         BackColor       =   &H8000000D&
         Height          =   1455
         Index           =   1
         Left            =   50
         TabIndex        =   23
         Top             =   1620
         Width           =   2780
      End
      Begin VB.Label selects 
         BackColor       =   &H8000000D&
         Height          =   1455
         Index           =   2
         Left            =   50
         TabIndex        =   24
         Top             =   3120
         Width           =   2780
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         Index           =   13
         Left            =   2800
         TabIndex        =   33
         Top             =   4800
         Width           =   1450
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         Index           =   12
         Left            =   2800
         TabIndex        =   32
         Top             =   4440
         Width           =   1450
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         Index           =   11
         Left            =   2800
         TabIndex        =   31
         Top             =   4080
         Width           =   1450
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         Index           =   10
         Left            =   2800
         TabIndex        =   30
         Top             =   3720
         Width           =   1450
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         Index           =   9
         Left            =   2800
         TabIndex        =   29
         Top             =   3360
         Width           =   1450
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         Index           =   8
         Left            =   2800
         TabIndex        =   28
         Top             =   3000
         Width           =   1450
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         Index           =   7
         Left            =   2800
         TabIndex        =   27
         Top             =   2640
         Width           =   1450
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         Index           =   6
         Left            =   2800
         TabIndex        =   26
         Top             =   2280
         Width           =   1450
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         Index           =   5
         Left            =   2800
         TabIndex        =   25
         Top             =   1920
         Width           =   1450
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Exit"
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
         Index           =   4
         Left            =   2800
         TabIndex        =   21
         Top             =   1560
         Width           =   1450
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Save"
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
         Index           =   3
         Left            =   2800
         TabIndex        =   20
         Top             =   1200
         Width           =   1450
      End
      Begin VB.Label menuItem 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Equip"
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
         Index           =   2
         Left            =   2800
         TabIndex        =   19
         Top             =   840
         Width           =   1450
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
         Height          =   375
         Index           =   1
         Left            =   2800
         TabIndex        =   18
         Top             =   480
         Width           =   1450
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
         Height          =   375
         Index           =   0
         Left            =   2800
         TabIndex        =   17
         Top             =   120
         Width           =   1450
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   0
      TabIndex        =   34
      Top             =   5760
      Width           =   4815
   End
End
Attribute VB_Name = "gameMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public inMenu, inMagic, inItem, inSelect, inWeapon As Boolean
Public isItem, isMagic, isWeapon As Boolean
Public menuSel, currMenu, currItem, currSelect, myMagic, myWeapon As Integer
Public itemType, magicType, weaponType As String
Public lastSel As Integer




Private Sub Form_Load()
    Dim a As Integer

    For a = 0 To 2
        If fightTeam(a) <> 0 Then
            team(a).Picture = people.guy((fightTeam(a) * 8) - 5).Picture
            myname(a).Caption = charNames(a)
            exper(a).Caption = "Exp. " & yourStats(fightTeam(a), EXP)
            hpoints(a).Caption = "HP: " & yourStats(fightTeam(a), CURHP) & "/" & yourStats(fightTeam(a), HP)
            mpoints(a).Caption = "MP: " & yourStats(fightTeam(a), CURMP) & "/" & yourStats(fightTeam(a), MP)
            lvl(a).Caption = "Lv. " & yourStats(fightTeam(a), LEVEL)
        Else
            team(a).Visible = False
            myname(a).Visible = False
            exper(a).Visible = False
            hpoints(a).Visible = False
            mpoints(a).Visible = False
            lvl(a).Visible = False
        End If
    Next a
    totalgil.Caption = "Gil: " & gil
    menuItem(0).BackColor = vbRed
    inMenu = True
    currMenu = 0
    magicType = "none"
    weaponType = "none"
End Sub




Private Sub RPG_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        gameScreen.Enabled = True
        Unload Me
    End If
    
    Label1.Caption = "inm:" & inMagic
        
    Dim tmpf As Integer
    Dim b As Integer
    Dim a As Integer
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyDown Then
        tmpf = 1
    End If
        If KeyCode = vbKeyUp Or KeyCode = vbKeyRight Then
        tmpf = -1
    End If
    
    
    
    If inMenu Then
        menuSel = 0
        menuItem(currMenu).BackColor = &H8000000D
        If (KeyCode = vbKeyUp Or KeyCode = vbKeyRight) And currMenu = 0 Then
            currMenu = 4
        Else
            currMenu = Abs(currMenu + tmpf) Mod 5
        End If
        menuItem(currMenu).BackColor = vbRed
    End If
    
    
'**************************************************
'*             SELECTING FROM MAIN MENU           *
'**************************************************
    If KeyCode = vbKeySpace Then
        If inMenu Then
            currItem = 0
            inMenu = False
            If currMenu = 0 Then
                currItem = 0
                inMenu = False
                inItem = True
                KeyCode = 0
                For a = 0 To 13
                    menuItem(a).Caption = itemInv(0, a) & "... " & itemInv(1, a)
                Next a
            End If
            If currMenu = 1 Then
                inMenu = False
                isMagic = True
                inSelect = True
                KeyCode = 0
            End If
            If currMenu = 2 Then
                inMenu = False
                isWeapon = True
                inSelect = True
                KeyCode = 0
            End If
        
        
        
        
        
        
            If currMenu = 4 Then
                gameScreen.Enabled = True
                Unload Me
            End If
        
        
        
        End If
    End If
    
'**************************************************
'*             SELECTING FROM ITEMS MENU          *
'**************************************************

    If inItem Then
        If KeyCode = vbKeySpace Then
            itemType = itemInv(0, menuSel)
            inSelect = True
            isItem = True
            inItem = False
            KeyCode = 0
        End If
        menuItem(currItem).BackColor = &H8000000D
        If menuSel + tmpf < 301 And tmpf > 0 Then
            If itemInv(0, menuSel + tmpf) <> "" Then
                menuSel = menuSel + tmpf
                If menuSel > 13 And currItem = 13 And tmpf = 1 Then
                    b = 13
                    For a = menuSel To menuSel - 13 Step -1
                        menuItem(b).Caption = itemInv(0, a) & "... " & itemInv(1, a)
                        b = b - 1
                    Next a
                End If
                If currItem < 13 And tmpf = 1 Then
                    currItem = currItem + 1
                End If
            End If
        Else
            If menuSel > 0 And tmpf < 0 Then
                If itemInv(0, menuSel + tmpf) <> "" Then
                    If menuSel > 0 And currItem = 0 Then
                        b = 0
                        For a = menuSel - 1 To menuSel + 12
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


    'in selection screen
    If inSelect Then
        If KeyCode = vbKeySpace And isItem And fightTeam(currSelect) <> 0 Then
            Dim itemID As Integer
            itemID = -1
            'remove used item from inventory
            For a = 0 To 299
                If itemInv(0, a) = itemType Then
                    itemID = a
                    If itemType <> "Phoenix" Then
                        If itemType = "Potion" And yourStats(fightTeam(currSelect), CURHP) < yourStats(fightTeam(currSelect), HP) Then
                            itemInv(1, a) = itemInv(1, a) - 1
                        End If
                        If itemType = "Ether" And yourStats(fightTeam(currSelect), CURMP) < yourStats(fightTeam(currSelect), MP) Then
                            itemInv(1, a) = itemInv(1, a) - 1
                        End If
                        If itemType = "Tent" And yourStats(fightTeam(currSelect), CURMP) < yourStats(fightTeam(currSelect), MP) Then
                            itemInv(1, a) = itemInv(1, a) - 1
                        End If
                    End If
                    menuItem(currItem).Caption = itemInv(0, a) & "... " & itemInv(1, a)
                    If itemInv(1, a) = 0 Then
                        For b = a To 299
                            Dim ts
                            Dim tn
                            ts = itemInv(0, b + 1)
                            tn = itemInv(1, b + 1)
                            itemInv(0, b + 1) = ""
                            itemInv(1, b + 1) = ""
                            itemInv(0, b) = ts
                            itemInv(1, b) = tn
                        Next b
                    End If
                    a = 300
                End If
            Next a
            If itemID <> -1 And itemType <> "Phoenix" Then
                If itemType = "Potion" And itemInv(1, itemID) >= 0 Then
                    'add to hp of selected player
                    If yourStats(fightTeam(currSelect), CURHP) >= yourStats(fightTeam(currSelect), HP) Then
                        itemType = "none"
                    End If
                    yourStats(fightTeam(currSelect), CURHP) = yourStats(fightTeam(currSelect), CURHP) + 50
                    If yourStats(fightTeam(currSelect), CURHP) > yourStats(fightTeam(currSelect), HP) Then
                        yourStats(fightTeam(currSelect), CURHP) = yourStats(fightTeam(currSelect), HP)
                    End If
                    hpoints(currSelect).Caption = "HP: " & yourStats(fightTeam(currSelect), CURHP) & "/" & yourStats(fightTeam(currSelect), HP)
                End If
                If itemType = "Ether" And itemInv(1, itemID) >= 0 And itemID <> -1 And yourStats(fightTeam(currSelect), CURMP) < yourStats(fightTeam(currSelect), MP) Then
                    'add to hp of selected player
                    If yourStats(fightTeam(currSelect), CURMP) >= yourStats(fightTeam(currSelect), MP) Then
                        itemType = "none"
                    End If
                    yourStats(fightTeam(currSelect), CURMP) = yourStats(fightTeam(currSelect), CURMP) + 50
                    If yourStats(fightTeam(currSelect), CURMP) > yourStats(fightTeam(currSelect), MP) Then
                        yourStats(fightTeam(currSelect), CURMP) = yourStats(fightTeam(currSelect), MP)
                    End If
                    mpoints(currSelect).Caption = "MP: " & yourStats(fightTeam(currSelect), CURMP) & "/" & yourStats(fightTeam(currSelect), MP)
                End If
                If itemType = "Tent" And itemInv(1, itemID) >= 0 Then
                    'add to hp of selected player
                    For a = 0 To 2
                        If fightTeam(a) <> 0 Then
                            yourStats(fightTeam(a), CURHP) = yourStats(fightTeam(a), HP)
                            yourStats(fightTeam(a), CURMP) = yourStats(fightTeam(a), MP)
                            hpoints(a).Caption = "HP: " & yourStats(fightTeam(a), CURHP) & "/" & yourStats(fightTeam(a), HP)
                            mpoints(a).Caption = "MP: " & yourStats(fightTeam(a), CURMP) & "/" & yourStats(fightTeam(a), MP)
                        End If
                    Next a
                End If
            
            
            End If
        End If
        
        
        
'*************************************************************
            'SETTING UP WEAPONS EQUIP
'*************************************************************
        If isWeapon And inWeapon = False And weaponType = "none" Then
            If KeyCode = vbKeySpace And fightTeam(currSelect) <> 0 Then
                For a = 0 To 2
                    selects(a).BackColor = &H8000000D
                Next a
                inWeapon = True
                KeyCode = 0
                inSelect = False
                myWeapon = fightTeam(currSelect)
                menuItem(2).BackColor = &H8000000D
                For a = 0 To 8
                    menuItem(a).Caption = weaponInv(0, myWeapon, a) & "... " & weaponInv(1, myWeapon, a)
                Next a
                
                menuItem(10).Caption = "Curr. Weapon:"
                menuItem(11).Caption = offeq(NAMED, myWeapon) & "... +" & offeq(POW, myWeapon)
                menuItem(12).Caption = "Curr. Shield:"
                menuItem(13).Caption = defeq(NAMED, myWeapon) & "... +" & defeq(POW, myWeapon)
                
            End If
        End If
        
       
        
'*************************************************************
            'SETTING UP MAGIC
'*************************************************************
       
        If isMagic And inMagic = False And magicType = "none" Then
            If KeyCode = vbKeySpace And fightTeam(currSelect) <> 0 Then
                For a = 0 To 2
                    selects(a).BackColor = &H8000000D
                Next a
                inMagic = True
                KeyCode = 0
                inSelect = False
                myMagic = fightTeam(currSelect)
                menuItem(1).BackColor = &H8000000D
                For a = 0 To 13
                    menuItem(a).Caption = magicInv(0, myMagic, a) & "... " & magicInv(1, myMagic, a)
                Next a
            End If
        End If
        
        selects(currSelect).BackColor = &H8000000D
        If (KeyCode = vbKeyUp Or KeyCode = vbKeyRight) And currSelect = 0 Then
            currSelect = 2
        Else
            currSelect = Abs(currSelect + tmpf) Mod 3
        End If
        selects(currSelect).BackColor = vbRed
        lastSel = currSelect
    End If
    
    
    'APPLY MAGIC SPELL IN MENU
    If inSelect And isMagic And magicType <> "none" And KeyCode = vbKeySpace And fightTeam(currSelect) <> 0 Then
        If magicType = "Cure" Then
            If yourStats(fightTeam(currSelect), CURHP) < yourStats(fightTeam(currSelect), HP) Then
                If yourStats(myMagic, CURMP) >= magicInv(1, myMagic, menuSel) Then
                    yourStats(myMagic, CURMP) = yourStats(myMagic, CURMP) - 6
                    yourStats(fightTeam(currSelect), CURHP) = yourStats(fightTeam(currSelect), CURHP) + (yourStats(fightTeam(currSelect), MAGPOWER) * (yourStats(fightTeam(currSelect), LEVEL) / 2))
                    mpoints(lastSel).Caption = "MP: " & yourStats(myMagic, CURMP) & "/" & yourStats(myMagic, MP)
                    If yourStats(fightTeam(currSelect), CURHP) > yourStats(fightTeam(currSelect), HP) Then
                        yourStats(fightTeam(currSelect), CURHP) = yourStats(fightTeam(currSelect), HP)
                    End If
                    hpoints(currSelect).Caption = "HP: " & yourStats(fightTeam(currSelect), CURHP) & "/" & yourStats(fightTeam(currSelect), HP)
                End If
            End If
        End If
    
        'add more spells...
    
    
    
    
    End If
    
    
    
    
    
    

    If inMagic And inSelect = False Then
        If KeyCode = vbKeySpace Then
            magicType = magicInv(0, myMagic, menuSel)
            If magicType = "Cure" Then
                inMagic = False
                inSelect = True
                For a = 0 To 2
                    selects(a).BackColor = &H8000000D
                Next a
                selects(currSelect).BackColor = vbRed
            Else
                magicType = "none"
            End If
        End If

        
        menuItem(currItem).BackColor = &H8000000D
        If menuSel + tmpf < 21 And tmpf > 0 Then
            If magicInv(0, myMagic, menuSel + tmpf) <> "" Then
                menuSel = menuSel + tmpf
                If menuSel > 13 And currItem = 13 And tmpf = 1 Then
                    b = 13
                    For a = menuSel To menuSel - 13 Step -1
                        menuItem(b).Caption = magicInv(0, myMagic, a) & "... " & magicInv(1, myMagic, a)
                        b = b - 1
                    Next a
                End If
                If currItem < 13 And tmpf = 1 Then
                    currItem = currItem + 1
                End If
            End If
        Else
            If menuSel > 0 And tmpf < 0 Then
                If magicInv(0, myMagic, menuSel + tmpf) <> "" Then
                    If menuSel > 0 And currItem = 0 Then
                        b = 0
                        For a = menuSel - 1 To menuSel + 12
                            menuItem(b).Caption = magicInv(0, myMagic, a) & "... " & magicInv(1, myMagic, a)
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

    
'*******************************************************************************
'       EQUIPPING EQUIPPING EQUIPPING
'**********************************************************************
    
    If inWeapon And isWeapon And weaponType = "none" And fightTeam(currSelect) <> 0 And KeyCode = vbKeySpace Then
         Dim tpow As Integer
         Dim tname As String
         If weaponInv(2, myWeapon, menuSel) = OFFENSE Then
            tpow = offeq(POW, myWeapon)
            tname = offeq(NAMED, myWeapon)
            offeq(POW, myWeapon) = weaponInv(POW, myWeapon, menuSel)
            offeq(NAMED, myWeapon) = weaponInv(NAMED, myWeapon, menuSel)
            weaponInv(POW, myWeapon, menuSel) = tpow
            weaponInv(NAMED, myWeapon, menuSel) = tname
            weaponInv(OFFDEF, myWeapon, menuSel) = OFFENSE
            menuItem(11).Caption = offeq(NAMED, myWeapon) & "... +" & offeq(POW, myWeapon)
            menuItem(currItem).Caption = weaponInv(0, myWeapon, menuSel) & "... " & weaponInv(1, myWeapon, menuSel)
         End If
         If weaponInv(2, myWeapon, menuSel) = DEFENSE Then
            tpow = defeq(POW, myWeapon)
            tname = defeq(NAMED, myWeapon)
            defeq(POW, myWeapon) = weaponInv(POW, myWeapon, menuSel)
            defeq(NAMED, myWeapon) = weaponInv(NAMED, myWeapon, menuSel)
            weaponInv(POW, myWeapon, menuSel) = tpow
            weaponInv(NAMED, myWeapon, menuSel) = tname
            weaponInv(OFFDEF, myWeapon, menuSel) = DEFENSE
            menuItem(13).Caption = defeq(NAMED, myWeapon) & "... +" & defeq(POW, myWeapon)
            menuItem(currItem).Caption = weaponInv(0, myWeapon, menuSel) & "... " & weaponInv(1, myWeapon, menuSel)
         End If
    
    End If





    If inWeapon And inSelect = False Then
        If KeyCode = vbKeySpace Then
            weaponType = "none"
        End If
        menuItem(currItem).BackColor = &H8000000D
        If menuSel + tmpf < 51 And tmpf > 0 Then
            If weaponInv(0, myWeapon, menuSel + tmpf) <> "" Then
                menuSel = menuSel + tmpf
                If menuSel > 8 And currItem = 8 And tmpf = 1 Then
                    b = 8
                    For a = menuSel To menuSel - 8 Step -1
                        menuItem(b).Caption = weaponInv(0, myWeapon, a) & "... " & weaponInv(1, myWeapon, a)
                        b = b - 1
                    Next a
                End If
                If currItem < 8 And tmpf = 1 Then
                    currItem = currItem + 1
                End If
            End If
        Else
            If menuSel > 0 And tmpf < 0 Then
                If weaponInv(0, myWeapon, menuSel + tmpf) <> "" Then
                    If menuSel > 0 And currItem = 0 Then
                        b = 0
                        For a = menuSel - 1 To menuSel + 7
                            menuItem(b).Caption = weaponInv(0, myWeapon, a) & "... " & weaponInv(1, myWeapon, a)
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









    'cancelling

    If KeyCode = vbKeyTab Then
        If inMenu Then
            gameScreen.Enabled = True
            Unload Me
        End If
        If inSelect And isItem Then
            inSelect = False
            For a = 0 To 2
                selects(a).BackColor = &H8000000D
            Next a
            inItem = True
            KeyCode = 0
        End If
        
        If inSelect And isMagic And inMagic = False And magicType = "none" Then
            inSelect = False
            inMenu = True
            isMagic = False
            For a = 0 To 2
                selects(a).BackColor = &H8000000D
            Next a
            KeyCode = 0
        End If
    
         If inSelect And isWeapon And inWeapon = False And weaponType = "none" Then
            inSelect = False
            inMenu = True
            isWeapon = False
            For a = 0 To 2
                selects(a).BackColor = &H8000000D
            Next a
            KeyCode = 0
        End If
        
        If inItem And KeyCode = vbKeyTab Then
            menuItem(currItem).BackColor = &H8000000D
            menuItem(0).Caption = "Item"
            menuItem(1).Caption = "Magic"
            menuItem(2).Caption = "Equip"
            menuItem(3).Caption = "Save"
            menuItem(4).Caption = "Exit"
            For a = 5 To 13
                menuItem(a).Caption = ""
            Next a
            menuItem(currMenu).BackColor = vbRed
            inItem = False
            isItem = False
            itemType = "none"
            inMenu = True
            KeyCode = 0
        End If
            
        If inMagic And magicType = "none" And isMagic Then
            inSelect = True
            inMagic = False
            For a = 0 To 2
                selects(a).BackColor = &H8000000D
            Next a
            selects(currSelect).BackColor = vbRed
            menuItem(currItem).BackColor = &H8000000D
            menuItem(0).Caption = "Item"
            menuItem(1).Caption = "Magic"
            menuItem(2).Caption = "Equip"
            menuItem(3).Caption = "Save"
            menuItem(4).Caption = "Exit"
            For a = 5 To 13
                menuItem(a).Caption = ""
            Next a
            menuItem(currMenu).BackColor = vbRed
            currItem = 0
        End If
    
        If inWeapon And weaponType = "none" And isWeapon Then
            inSelect = True
            inWeapon = False
            For a = 0 To 2
                selects(a).BackColor = &H8000000D
            Next a
            selects(currSelect).BackColor = vbRed
            menuItem(currItem).BackColor = &H8000000D
            menuItem(0).Caption = "Item"
            menuItem(1).Caption = "Magic"
            menuItem(2).Caption = "Equip"
            menuItem(3).Caption = "Save"
            menuItem(4).Caption = "Exit"
            For a = 5 To 13
                menuItem(a).Caption = ""
            Next a
            menuItem(currMenu).BackColor = vbRed
            currItem = 0
        End If
    
        If inSelect And isMagic And magicType <> "none" Then
            inSelect = False
            inMagic = True
            For a = 0 To 2
                selects(a).BackColor = &H8000000D
            Next a
            selects(currSelect).BackColor = vbRed
            magicType = "none"
        End If
        
        If inSelect And isWeapon And weaponType <> "none" Then
            inSelect = False
            isWeapon = False
            For a = 0 To 2
                selects(a).BackColor = &H8000000D
            Next a
            selects(myWeapon).BackColor = vbRed
            weaponType = "none"
        End If
    
    
    
    End If




    
    
    
    
End Sub

Private Sub Timer1_Timer()
'Label1.Caption = " inm: " & inWeapon & " ism: " & isWeapon & " ins: " & inSelect & " mtype: " & weaponType

End Sub
