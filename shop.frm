VERSION 5.00
Begin VB.Form shop 
   Caption         =   "Shop"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox RPG 
      BackColor       =   &H00800000&
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   4400
      Left            =   240
      ScaleHeight     =   4335
      ScaleWidth      =   4335
      TabIndex        =   1
      Top             =   240
      Width           =   4400
      Begin VB.Label yourGil 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label yourAmount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   495
         Left            =   3120
         TabIndex        =   30
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
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
         Height          =   495
         Left            =   3120
         TabIndex        =   29
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label yourPrice 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   495
         Left            =   3120
         TabIndex        =   28
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Price:"
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
         Height          =   495
         Left            =   3120
         TabIndex        =   27
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label youHave 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   495
         Left            =   3240
         TabIndex        =   26
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "You Have:"
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
         Height          =   495
         Left            =   3120
         TabIndex        =   25
         Top             =   480
         Width           =   1215
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000009&
         Height          =   4095
         Left            =   3000
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label menuPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   10
         Left            =   1920
         TabIndex        =   24
         Top             =   3960
         Width           =   1005
      End
      Begin VB.Label menuPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   9
         Left            =   1920
         TabIndex        =   23
         Top             =   3600
         Width           =   1005
      End
      Begin VB.Label menuPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   8
         Left            =   1920
         TabIndex        =   22
         Top             =   3240
         Width           =   1005
      End
      Begin VB.Label menuPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   7
         Left            =   1920
         TabIndex        =   21
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Label menuPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   6
         Left            =   1920
         TabIndex        =   20
         Top             =   2520
         Width           =   1005
      End
      Begin VB.Label menuPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   19
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label menuPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   4
         Left            =   1920
         TabIndex        =   18
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label menuPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   17
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label menuPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   2
         Left            =   1920
         TabIndex        =   16
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label menuPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   15
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label menuPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label menuItem 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   10
         Left            =   45
         TabIndex        =   13
         Top             =   3960
         Width           =   1875
      End
      Begin VB.Label menuItem 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   9
         Left            =   45
         TabIndex        =   12
         Top             =   3600
         Width           =   1875
      End
      Begin VB.Label menuItem 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   8
         Left            =   45
         TabIndex        =   11
         Top             =   3240
         Width           =   1875
      End
      Begin VB.Label menuItem 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   7
         Left            =   45
         TabIndex        =   10
         Top             =   2880
         Width           =   1875
      End
      Begin VB.Label menuItem 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   6
         Left            =   45
         TabIndex        =   9
         Top             =   2520
         Width           =   1875
      End
      Begin VB.Label menuItem 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   5
         Left            =   45
         TabIndex        =   8
         Top             =   2160
         Width           =   1875
      End
      Begin VB.Label menuItem 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   4
         Left            =   45
         TabIndex        =   7
         Top             =   1800
         Width           =   1875
      End
      Begin VB.Label menuItem 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   3
         Left            =   45
         TabIndex        =   6
         Top             =   1440
         Width           =   1875
      End
      Begin VB.Label menuItem 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   2
         Left            =   45
         TabIndex        =   5
         Top             =   1080
         Width           =   1875
      End
      Begin VB.Label menuItem 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   1
         Left            =   45
         TabIndex        =   4
         Top             =   720
         Width           =   1875
      End
      Begin VB.Label menuItem 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   0
         Left            =   45
         TabIndex        =   3
         Top             =   360
         Width           =   1875
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000009&
         Height          =   4095
         Left            =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label buysellexit 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Buy"
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
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   735
      End
      Begin VB.Label buysellexit 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
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
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   32
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Label dialogue 
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Welcome to my store."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   4680
      Width           =   4395
   End
End
Attribute VB_Name = "shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public inBuySell, inBuyMenu, inSellMenu, inChooseNum As Boolean
Public currBS, menuSel, currItem As Integer
Public getAmt As Integer
Public itemID As Integer




Private Sub Form_Load()
inBuySell = True
currBS = 0
menuSel = 0
getAmt = 1
buysellexit(currBS).BackColor = vbRed
yourGil.Caption = "Gil: " & gil
End Sub



Private Sub RPG_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        gameScreen.Enabled = True
        Unload Me
    End If
    
    Dim tmpf As Integer
    
    
    If KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
        tmpf = 1
    End If
    If KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
        tmpf = -1
    End If
    
'***************************************************
'*      in BUY SELL EXIT menu  BUY SELL EXIT       *
'***************************************************
    
    If inBuySell Then
        buysellexit(currBS).BackColor = &H8000000D
        If currBS = 0 And tmpf = -1 Then
            currBS = 1
        Else
            currBS = (currBS + tmpf) Mod 2
        End If
        buysellexit(currBS).BackColor = vbRed
        
        'move onto next menu
        If KeyCode = vbKeySpace Then
            If currBS = 0 Then
                inBuyMenu = True
                inBuySell = False
                KeyCode = 0
                'sID------v
                Call getStore(0)
                For a = 0 To 10
                    'sID------------------------------v
                    menuItem(a).Caption = storeInv(0, 0, a)
                    menuPrice(a).Caption = storeInv(1, 0, a)
                Next a
                menuItem(0).BackColor = vbRed
                menuPrice(0).BackColor = vbRed
            End If

            If currBS = 1 Then
                gameScreen.Enabled = True
                gameScreen.dialogue = "Clerk:  Thanks!"
                Unload Me
            End If
        End If
    End If
    
    
    
    
'***************************************************
'*      in BUY BUY  BUY  menu  BUY BUY BUY BUY     *
'***************************************************

    If inBuyMenu Then
'***************************************************
'*      CHOOSING AMOUNT TO BUY AMOUNT CHOICE       *
'***************************************************
        If KeyCode = vbKeySpace Then
            dialogue.Caption = "Choose your amount, please."
            inBuyMenu = False
            inChooseNum = True
            KeyCode = 0
        Else
            menuItem(currItem).BackColor = &H8000000D
            menuPrice(currItem).BackColor = &H8000000D
            If menuSel + tmpf < 301 And tmpf > 0 Then
                'sID----------v
                If storeInv(0, 0, menuSel + tmpf) <> "" Then
                    menuSel = menuSel + tmpf
                    If menuSel > 10 And currItem = 10 And tmpf = 1 Then
                        b = 10
                        For a = menuSel To menuSel - 10 Step -1
                            'must figure out sID--------------v
                            menuItem(b).Caption = storeInv(0, 0, a)
                            menuPrice(b).Caption = storeInv(1, 0, a)
                            b = b - 1
                        Next a
                    End If
                    If currItem < 10 And tmpf = 1 Then
                        currItem = currItem + 1
                    End If
                End If
            Else
                If menuSel > 0 And tmpf < 0 Then
                    'sID-----------v
                    If storeInv(0, 0, menuSel + tmpf) <> "" Then
                        If menuSel > 0 And currItem = 0 Then
                            b = 0
                            For a = menuSel - 1 To menuSel + 9
                                'sID------------------------------v
                                menuItem(b).Caption = storeInv(0, 0, a)
                                menuPrice(b).Caption = storeInv(1, 0, a)
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
            menuPrice(currItem).BackColor = vbRed
        End If
    End If
    
    
    
    
'***************************************************
'*      CHOOSING AMOUNT TO BUY AMOUNT CHOICE       *
'***************************************************
    If inChooseNum Then
        Dim totprice As Integer
        If KeyCode = vbKeySpace Then
            yourAmount.Caption = getAmt
            totprice = getAmt * storeInv(1, 0, menuSel)
            If totprice > gil Then
                getAmt = getAmt - 1
                totprice = getAmt * storeInv(1, 0, menuSel)
            End If
            yourAmount.Caption = getAmt
            yourPrice.Caption = totprice
                        
            gil = gil - totprice
            
            yourGil.Caption = "Gil: " & gil
            itemInv(1, itemID) = itemInv(1, itemID) + getAmt
            'sID-----------------------------v
            itemInv(0, itemID) = storeInv(0, 0, menuSel)
            getAmt = 1
            totprice = storeInv(1, 0, menuSel)
            yourAmount.Caption = getAmt
            yourPrice.Caption = totprice
            For a = 0 To 300
                If storeInv(0, 0, menuSel) = itemInv(0, a) Then
                    youHave.Caption = itemInv(1, a)
                    a = 301
                End If
            Next a
            dialogue = "THANKS!"
        
        
        Else
            youHave.Caption = "0"
            Dim yours As Integer
            Dim max As Integer
            max = 99
            For a = 0 To 300
                If storeInv(0, 0, menuSel) = itemInv(0, a) Then
                    itemID = a
                    youHave.Caption = itemInv(1, a)
                    max = 99 - itemInv(1, a)
                    a = 305
                End If
            Next a
            If a <> 306 Then
                For a = 0 To 300
                    If itemInv(0, a) = "" Then
                        itemID = a
                        a = 301
                    End If
                Next a
            End If
            getAmt = getAmt - tmpf
            If getAmt < 1 Then
                getAmt = 1
            End If
            If getAmt > max Then
                getAmt = max
            End If
            yourAmount.Caption = getAmt
            totprice = getAmt * storeInv(1, 0, menuSel)
            If totprice > gil Then
                getAmt = getAmt - 1
                totprice = getAmt * storeInv(1, 0, menuSel)
            End If
            yourAmount.Caption = getAmt
            yourPrice.Caption = totprice
        End If
    End If

'***************************************************
'*      CANCEL OUT OF A MENU CANCELLING            *
'***************************************************
    If KeyCode = vbKeyTab Then
        If inBuySell Then
            gameScreen.Enabled = True
            Unload Me
        End If
        If inBuyMenu Or inSellMenu Then
            inBuyMenu = False
            inSellMenu = False
            inBuySell = True
            menuSel = 0
            currItem = 0
            For a = 0 To 10
                menuItem(a).Caption = ""
                menuItem(a).BackColor = &H8000000D
                menuPrice(a).Caption = ""
                menuPrice(a).BackColor = &H8000000D
            Next a
        End If
        If inChooseNum Then
            yourPrice.Caption = ""
            yourAmount.Caption = ""
            youHave.Caption = ""
            inBuyMenu = True
            inChooseNum = False
            getAmt = 1
            dialogue.Caption = "Welcome to my store!"
        End If
    End If
End Sub
