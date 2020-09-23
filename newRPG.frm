VERSION 5.00
Begin VB.Form gameScreen 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   1785
   ClientTop       =   1140
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer movement 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5640
      Top             =   2160
   End
   Begin VB.PictureBox RPG 
      Height          =   4400
      Left            =   240
      ScaleHeight     =   4335
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   240
      Width           =   4400
      Begin VB.Image man 
         Height          =   480
         Left            =   480
         Picture         =   "newRPG.frx":0000
         Top             =   0
         Width           =   360
      End
      Begin VB.Image tile 
         Height          =   495
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   213
      Left            =   10080
      Picture         =   "newRPG.frx":0407
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   214
      Left            =   10440
      Picture         =   "newRPG.frx":07BD
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   205
      Left            =   10080
      Picture         =   "newRPG.frx":0B73
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   206
      Left            =   10440
      Picture         =   "newRPG.frx":0EB5
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   207
      Left            =   10080
      Picture         =   "newRPG.frx":11F7
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   208
      Left            =   10440
      Picture         =   "newRPG.frx":1539
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   209
      Left            =   10080
      Picture         =   "newRPG.frx":187B
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   210
      Left            =   10440
      Picture         =   "newRPG.frx":1BBD
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   211
      Left            =   10080
      Picture         =   "newRPG.frx":1EFF
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   212
      Left            =   10440
      Picture         =   "newRPG.frx":2241
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   251
      Left            =   4680
      Picture         =   "newRPG.frx":2583
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   250
      Left            =   4320
      Picture         =   "newRPG.frx":28C5
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   249
      Left            =   4680
      Picture         =   "newRPG.frx":2C07
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   248
      Left            =   4320
      Picture         =   "newRPG.frx":2F49
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   247
      Left            =   8040
      Picture         =   "newRPG.frx":328B
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   246
      Left            =   7800
      Picture         =   "newRPG.frx":35CD
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   245
      Left            =   8280
      Picture         =   "newRPG.frx":390F
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   244
      Left            =   7320
      Picture         =   "newRPG.frx":3C51
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   243
      Left            =   7320
      Picture         =   "newRPG.frx":3F93
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   242
      Left            =   7320
      Picture         =   "newRPG.frx":42D5
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   241
      Left            =   7320
      Picture         =   "newRPG.frx":4617
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   240
      Left            =   6840
      Picture         =   "newRPG.frx":4959
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   239
      Left            =   5760
      Picture         =   "newRPG.frx":4C9B
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   238
      Left            =   6120
      Picture         =   "newRPG.frx":4FDD
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   237
      Left            =   5760
      Picture         =   "newRPG.frx":531F
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   236
      Left            =   5760
      Picture         =   "newRPG.frx":5661
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   235
      Left            =   6120
      Picture         =   "newRPG.frx":59A3
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   234
      Left            =   6120
      Picture         =   "newRPG.frx":5CE5
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   233
      Left            =   6480
      Picture         =   "newRPG.frx":6067
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   232
      Left            =   6480
      Picture         =   "newRPG.frx":63A9
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   231
      Left            =   6480
      Picture         =   "newRPG.frx":66EB
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   230
      Left            =   11640
      Picture         =   "newRPG.frx":6A2D
      Stretch         =   -1  'True
      Top             =   480
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   229
      Left            =   11160
      Picture         =   "newRPG.frx":6DE8
      Stretch         =   -1  'True
      Top             =   480
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   228
      Left            =   11640
      Picture         =   "newRPG.frx":71A4
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   227
      Left            =   11160
      Picture         =   "newRPG.frx":755A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   226
      Left            =   11640
      Picture         =   "newRPG.frx":7910
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   225
      Left            =   11160
      Picture         =   "newRPG.frx":7CC6
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   224
      Left            =   11640
      Picture         =   "newRPG.frx":807C
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   223
      Left            =   11160
      Picture         =   "newRPG.frx":83BE
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   222
      Left            =   11640
      Picture         =   "newRPG.frx":8700
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   221
      Left            =   11160
      Picture         =   "newRPG.frx":8A42
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   220
      Left            =   11640
      Picture         =   "newRPG.frx":8D84
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   219
      Left            =   11160
      Picture         =   "newRPG.frx":913A
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   218
      Left            =   11640
      Picture         =   "newRPG.frx":94F0
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   217
      Left            =   11160
      Picture         =   "newRPG.frx":98A6
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   216
      Left            =   11640
      Picture         =   "newRPG.frx":9C5C
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   215
      Left            =   11160
      Picture         =   "newRPG.frx":A012
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   203
      Left            =   10080
      Picture         =   "newRPG.frx":A3C8
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   204
      Left            =   10560
      Picture         =   "newRPG.frx":A77E
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   201
      Left            =   10080
      Picture         =   "newRPG.frx":AB34
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   202
      Left            =   10560
      Picture         =   "newRPG.frx":AEEA
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   199
      Left            =   10080
      Picture         =   "newRPG.frx":B2A0
      Stretch         =   -1  'True
      Top             =   480
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   200
      Left            =   10560
      Picture         =   "newRPG.frx":B5E2
      Stretch         =   -1  'True
      Top             =   480
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   27
      Left            =   8280
      Picture         =   "newRPG.frx":B924
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   26
      Left            =   7800
      Picture         =   "newRPG.frx":BCDA
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   194
      Left            =   7560
      Picture         =   "newRPG.frx":C090
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Y"
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "X"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   25
      Left            =   8280
      Picture         =   "newRPG.frx":C3D2
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   24
      Left            =   7800
      Picture         =   "newRPG.frx":C788
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   23
      Left            =   8280
      Picture         =   "newRPG.frx":CB3E
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   22
      Left            =   7800
      Picture         =   "newRPG.frx":CE80
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label dialogue 
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   6
      Top             =   4680
      Width           =   4395
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   29
      Left            =   8280
      Picture         =   "newRPG.frx":D1C2
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   28
      Left            =   7800
      Picture         =   "newRPG.frx":D578
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   191
      Left            =   7200
      Picture         =   "newRPG.frx":D92E
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   192
      Left            =   7560
      Picture         =   "newRPG.frx":DC70
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   193
      Left            =   7200
      Picture         =   "newRPG.frx":DFB2
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   195
      Left            =   7920
      Picture         =   "newRPG.frx":E2F4
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   196
      Left            =   8280
      Picture         =   "newRPG.frx":E636
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   197
      Left            =   7920
      Picture         =   "newRPG.frx":E978
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   198
      Left            =   8280
      Picture         =   "newRPG.frx":ECBA
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   49
      Left            =   9600
      Picture         =   "newRPG.frx":EFFC
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   50
      Left            =   8880
      Picture         =   "newRPG.frx":F33E
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   51
      Left            =   9600
      Picture         =   "newRPG.frx":F680
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   52
      Left            =   8880
      Picture         =   "newRPG.frx":F9C2
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   53
      Left            =   8880
      Picture         =   "newRPG.frx":FD04
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   54
      Left            =   9240
      Picture         =   "newRPG.frx":10046
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   55
      Left            =   9600
      Picture         =   "newRPG.frx":10388
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   56
      Left            =   9240
      Picture         =   "newRPG.frx":106CA
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   57
      Left            =   10440
      Picture         =   "newRPG.frx":10A0C
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   0
      Left            =   9240
      Picture         =   "newRPG.frx":10D4E
      Stretch         =   -1  'True
      Top             =   480
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   41
      Left            =   9240
      Picture         =   "newRPG.frx":11090
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   42
      Left            =   9600
      Picture         =   "newRPG.frx":113D2
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   43
      Left            =   9240
      Picture         =   "newRPG.frx":11714
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   44
      Left            =   8880
      Picture         =   "newRPG.frx":11A56
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   45
      Left            =   8880
      Picture         =   "newRPG.frx":11D98
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   46
      Left            =   9600
      Picture         =   "newRPG.frx":120DA
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   47
      Left            =   8880
      Picture         =   "newRPG.frx":1241C
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   48
      Left            =   9600
      Picture         =   "newRPG.frx":1275E
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   58
      Left            =   9240
      Picture         =   "newRPG.frx":12AA0
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   59
      Left            =   8880
      Picture         =   "newRPG.frx":12DE2
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   60
      Left            =   9240
      Picture         =   "newRPG.frx":13124
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   61
      Left            =   9600
      Picture         =   "newRPG.frx":13466
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   62
      Left            =   9600
      Picture         =   "newRPG.frx":137A8
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   63
      Left            =   8880
      Picture         =   "newRPG.frx":13AEA
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   64
      Left            =   8880
      Picture         =   "newRPG.frx":13E2C
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   65
      Left            =   9600
      Picture         =   "newRPG.frx":1416E
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   1
      Left            =   9240
      Picture         =   "newRPG.frx":144B0
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   32
      Left            =   9600
      Picture         =   "newRPG.frx":147F2
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   33
      Left            =   8880
      Picture         =   "newRPG.frx":14B34
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   34
      Left            =   9600
      Picture         =   "newRPG.frx":14E76
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   35
      Left            =   8880
      Picture         =   "newRPG.frx":151B8
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   36
      Left            =   8880
      Picture         =   "newRPG.frx":154FA
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   37
      Left            =   9240
      Picture         =   "newRPG.frx":1583C
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   38
      Left            =   9600
      Picture         =   "newRPG.frx":15B7E
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   39
      Left            =   9240
      Picture         =   "newRPG.frx":15EC0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   190
      Left            =   6480
      Picture         =   "newRPG.frx":16202
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   189
      Left            =   5760
      Picture         =   "newRPG.frx":16544
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   188
      Left            =   5760
      Picture         =   "newRPG.frx":16886
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   187
      Left            =   5760
      Picture         =   "newRPG.frx":16BC8
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   186
      Left            =   6480
      Picture         =   "newRPG.frx":16F0A
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   185
      Left            =   5760
      Picture         =   "newRPG.frx":1724C
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   184
      Left            =   5400
      Picture         =   "newRPG.frx":1758E
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   183
      Left            =   6840
      Picture         =   "newRPG.frx":178D0
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   182
      Left            =   6480
      Picture         =   "newRPG.frx":17C12
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   181
      Left            =   6120
      Picture         =   "newRPG.frx":17F54
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   180
      Left            =   6120
      Picture         =   "newRPG.frx":18296
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   179
      Left            =   6480
      Picture         =   "newRPG.frx":185D8
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   178
      Left            =   5400
      Picture         =   "newRPG.frx":1891A
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   177
      Left            =   6840
      Picture         =   "newRPG.frx":18C5C
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   176
      Left            =   5760
      Picture         =   "newRPG.frx":18F9E
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   175
      Left            =   5760
      Picture         =   "newRPG.frx":192E0
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   174
      Left            =   6480
      Picture         =   "newRPG.frx":19622
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   173
      Left            =   6120
      Picture         =   "newRPG.frx":19964
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   172
      Left            =   6480
      Picture         =   "newRPG.frx":19CE6
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   171
      Left            =   6120
      Picture         =   "newRPG.frx":1A028
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   170
      Left            =   6960
      Picture         =   "newRPG.frx":1A36A
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   169
      Left            =   6960
      Picture         =   "newRPG.frx":1A6AC
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   168
      Left            =   7320
      Picture         =   "newRPG.frx":1A9EE
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   167
      Left            =   6960
      Picture         =   "newRPG.frx":1AD30
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   166
      Left            =   7320
      Picture         =   "newRPG.frx":1B072
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   165
      Left            =   7320
      Picture         =   "newRPG.frx":1B3B4
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   164
      Left            =   7680
      Picture         =   "newRPG.frx":1B6F6
      Stretch         =   -1  'True
      Top             =   9720
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   163
      Left            =   7680
      Picture         =   "newRPG.frx":1BA38
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   162
      Left            =   7680
      Picture         =   "newRPG.frx":1BD7A
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   161
      Left            =   7680
      Picture         =   "newRPG.frx":1C0BC
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   160
      Left            =   7680
      Picture         =   "newRPG.frx":1C3FE
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   159
      Left            =   9120
      Picture         =   "newRPG.frx":1C740
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   158
      Left            =   9120
      Picture         =   "newRPG.frx":1CA82
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   157
      Left            =   9120
      Picture         =   "newRPG.frx":1CDC4
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   156
      Left            =   9120
      Picture         =   "newRPG.frx":1D106
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   155
      Left            =   8040
      Picture         =   "newRPG.frx":1D448
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   154
      Left            =   8040
      Picture         =   "newRPG.frx":1D78A
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   153
      Left            =   8040
      Picture         =   "newRPG.frx":1DACC
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   152
      Left            =   8040
      Picture         =   "newRPG.frx":1DE0E
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   151
      Left            =   8760
      Picture         =   "newRPG.frx":1E150
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   150
      Left            =   8400
      Picture         =   "newRPG.frx":1E492
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   149
      Left            =   8400
      Picture         =   "newRPG.frx":1E7D4
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   148
      Left            =   8760
      Picture         =   "newRPG.frx":1EB16
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   147
      Left            =   8400
      Picture         =   "newRPG.frx":1EE58
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   146
      Left            =   8760
      Picture         =   "newRPG.frx":1F19A
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   145
      Left            =   8400
      Picture         =   "newRPG.frx":1F4DC
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   8
      Left            =   8760
      Picture         =   "newRPG.frx":1F81E
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   144
      Left            =   5280
      Picture         =   "newRPG.frx":1FB60
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   143
      Left            =   5640
      Picture         =   "newRPG.frx":1FF16
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   142
      Left            =   5280
      Picture         =   "newRPG.frx":202CC
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   141
      Left            =   5280
      Picture         =   "newRPG.frx":20682
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   140
      Left            =   5640
      Picture         =   "newRPG.frx":20A38
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   139
      Left            =   6000
      Picture         =   "newRPG.frx":20DEE
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   138
      Left            =   6000
      Picture         =   "newRPG.frx":211A4
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   137
      Left            =   6000
      Picture         =   "newRPG.frx":2155A
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   136
      Left            =   3840
      Picture         =   "newRPG.frx":21910
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   135
      Left            =   3840
      Picture         =   "newRPG.frx":21CC6
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   134
      Left            =   4920
      Picture         =   "newRPG.frx":2207C
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   133
      Left            =   4920
      Picture         =   "newRPG.frx":22432
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   132
      Left            =   4920
      Picture         =   "newRPG.frx":227E8
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   131
      Left            =   4560
      Picture         =   "newRPG.frx":22B9E
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   130
      Left            =   4560
      Picture         =   "newRPG.frx":22F54
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   129
      Left            =   4200
      Picture         =   "newRPG.frx":2330A
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   128
      Left            =   4200
      Picture         =   "newRPG.frx":236C0
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   127
      Left            =   4560
      Picture         =   "newRPG.frx":23A76
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   126
      Left            =   4200
      Picture         =   "newRPG.frx":23E2C
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   125
      Left            =   3480
      Picture         =   "newRPG.frx":241E2
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   124
      Left            =   3840
      Picture         =   "newRPG.frx":24524
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   123
      Left            =   3480
      Picture         =   "newRPG.frx":24866
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   122
      Left            =   3840
      Picture         =   "newRPG.frx":24BA8
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   121
      Left            =   3120
      Picture         =   "newRPG.frx":24EEA
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   120
      Left            =   3120
      Picture         =   "newRPG.frx":2522C
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   119
      Left            =   2760
      Picture         =   "newRPG.frx":2556E
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   118
      Left            =   2760
      Picture         =   "newRPG.frx":258B0
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   117
      Left            =   1680
      Picture         =   "newRPG.frx":25BF2
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   116
      Left            =   1680
      Picture         =   "newRPG.frx":25F34
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   115
      Left            =   960
      Picture         =   "newRPG.frx":26276
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   114
      Left            =   600
      Picture         =   "newRPG.frx":265B8
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   113
      Left            =   3480
      Picture         =   "newRPG.frx":268FA
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   112
      Left            =   3120
      Picture         =   "newRPG.frx":26C3C
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   111
      Left            =   3120
      Picture         =   "newRPG.frx":26F7E
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   110
      Left            =   3480
      Picture         =   "newRPG.frx":272C0
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   109
      Left            =   600
      Picture         =   "newRPG.frx":27602
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   108
      Left            =   960
      Picture         =   "newRPG.frx":279B8
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   107
      Left            =   600
      Picture         =   "newRPG.frx":27D6E
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   106
      Left            =   600
      Picture         =   "newRPG.frx":28124
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   105
      Left            =   960
      Picture         =   "newRPG.frx":284DA
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   104
      Left            =   960
      Picture         =   "newRPG.frx":28890
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   103
      Left            =   1680
      Picture         =   "newRPG.frx":28C46
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   102
      Left            =   1680
      Picture         =   "newRPG.frx":29348
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   101
      Left            =   2760
      Picture         =   "newRPG.frx":29A4A
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   100
      Left            =   2760
      Picture         =   "newRPG.frx":29E00
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   91
      Left            =   1320
      Picture         =   "newRPG.frx":2A1B6
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   92
      Left            =   1320
      Picture         =   "newRPG.frx":2A56C
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   93
      Left            =   1320
      Picture         =   "newRPG.frx":2A922
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   94
      Left            =   1320
      Picture         =   "newRPG.frx":2ACD8
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   95
      Left            =   2040
      Picture         =   "newRPG.frx":2B08E
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   96
      Left            =   2040
      Picture         =   "newRPG.frx":2B444
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   97
      Left            =   2400
      Picture         =   "newRPG.frx":2B7FA
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   98
      Left            =   2400
      Picture         =   "newRPG.frx":2BBB0
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   99
      Left            =   2400
      Picture         =   "newRPG.frx":2BF66
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   89
      Left            =   12960
      Picture         =   "newRPG.frx":2C31C
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   90
      Left            =   12600
      Picture         =   "newRPG.frx":2DD9E
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "countertop"
      Height          =   255
      Left            =   11520
      TabIndex        =   5
      Top             =   5280
      Width           =   855
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   88
      Left            =   11760
      Picture         =   "newRPG.frx":2F820
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   87
      Left            =   11160
      Picture         =   "newRPG.frx":312A2
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   86
      Left            =   11160
      Picture         =   "newRPG.frx":32D24
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   85
      Left            =   9480
      Picture         =   "newRPG.frx":345D6
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   84
      Left            =   9840
      Picture         =   "newRPG.frx":35E88
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   7
      Left            =   8760
      Picture         =   "newRPG.frx":3773A
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   75
      Left            =   8760
      Picture         =   "newRPG.frx":391BC
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   76
      Left            =   9480
      Picture         =   "newRPG.frx":3AB26
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   77
      Left            =   9480
      Picture         =   "newRPG.frx":3C490
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   78
      Left            =   8760
      Picture         =   "newRPG.frx":3DDFA
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   79
      Left            =   9480
      Picture         =   "newRPG.frx":3F764
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   80
      Left            =   8760
      Picture         =   "newRPG.frx":410CE
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   81
      Left            =   9120
      Picture         =   "newRPG.frx":42A38
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   82
      Left            =   9120
      Picture         =   "newRPG.frx":442EA
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   83
      Left            =   9120
      Picture         =   "newRPG.frx":45B9C
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   74
      Left            =   9600
      Picture         =   "newRPG.frx":4744E
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   73
      Left            =   8880
      Picture         =   "newRPG.frx":47790
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   72
      Left            =   9600
      Picture         =   "newRPG.frx":47AD2
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   71
      Left            =   8880
      Picture         =   "newRPG.frx":47E14
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   70
      Left            =   8880
      Picture         =   "newRPG.frx":48156
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   69
      Left            =   9240
      Picture         =   "newRPG.frx":48498
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   68
      Left            =   9600
      Picture         =   "newRPG.frx":487DA
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   67
      Left            =   9240
      Picture         =   "newRPG.frx":48B1C
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   66
      Left            =   9240
      Picture         =   "newRPG.frx":48E5E
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   40
      Left            =   8040
      Picture         =   "newRPG.frx":491A0
      Stretch         =   -1  'True
      Top             =   9720
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   30
      Left            =   11760
      Picture         =   "newRPG.frx":49556
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   21
      Left            =   8280
      Picture         =   "newRPG.frx":4AFD8
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   20
      Left            =   7800
      Picture         =   "newRPG.frx":4B31A
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   19
      Left            =   8280
      Picture         =   "newRPG.frx":4B65C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   18
      Left            =   7800
      Picture         =   "newRPG.frx":4B99E
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   17
      Left            =   8280
      Picture         =   "newRPG.frx":4BCE0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   16
      Left            =   7800
      Picture         =   "newRPG.frx":4C096
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   15
      Left            =   8280
      Picture         =   "newRPG.frx":4C44C
      Stretch         =   -1  'True
      Top             =   480
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   14
      Left            =   7800
      Picture         =   "newRPG.frx":4C78E
      Stretch         =   -1  'True
      Top             =   480
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   13
      Left            =   13200
      Picture         =   "newRPG.frx":4CAD0
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   11
      Left            =   2400
      Picture         =   "newRPG.frx":4CE12
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   12
      Left            =   2040
      Picture         =   "newRPG.frx":4D1C8
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   6
      Left            =   11280
      Picture         =   "newRPG.frx":4D57E
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   9
      Left            =   12720
      Picture         =   "newRPG.frx":4F000
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   "block"
      Height          =   255
      Left            =   12600
      TabIndex        =   4
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "town"
      Height          =   255
      Left            =   13560
      TabIndex        =   3
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   10
      Left            =   13560
      Picture         =   "newRPG.frx":50A82
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image character 
      Height          =   375
      Left            =   10920
      Picture         =   "newRPG.frx":52504
      Stretch         =   -1  'True
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   5
      Left            =   10560
      Picture         =   "newRPG.frx":53F86
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   4
      Left            =   10440
      Picture         =   "newRPG.frx":55A08
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   3
      Left            =   9240
      Picture         =   "newRPG.frx":5748A
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   2
      Left            =   9240
      Picture         =   "newRPG.frx":577CC
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "roof"
      Height          =   255
      Left            =   10440
      TabIndex        =   2
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "castle"
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   6960
      Width           =   495
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   31
      Left            =   10320
      Picture         =   "newRPG.frx":57B0E
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Menu quitg 
      Caption         =   "exit"
   End
End
Attribute VB_Name = "gameScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public inHotSpot, inDoor, isTalking As Boolean
Public lastx, lasty, lastdoorx, lastdoory, cntr, eraseTalk, walkctr As Integer



Private Sub Form_Unload(Cancel As Integer)
    End
End Sub





Private Sub Form_Load()

talked = False
'talked = True


charNames(0) = "Samweis"
Dim a As Integer


'load enemy stats into arrays
For a = 1 To 3
    Call getStats(a)
Next a


Call setYourInitStats(1)
Call loadMapInfo
Call resetTreasure
Call setTreasure
Dim tmp As Integer
party(0) = 1
fightTeam(0) = 1
Call checkLevel(1)



lastdoorx = 3
lastdoory = 8
lastx = 4
lasty = 4
charX = 14
charY = 16
mapX = 1
mapY = 7
eraseTalk = 0
walkctr = 0
inDoor = True
inHotSpot = True
movement.Enabled = True

'reset maps/treasure for new games!
    For a = 1 To 5
        Call undoPerform(a)
    Next a
    

    For row = 0 To 19
        For col = 0 To 19
            tmp = (row * 20) + col
            If tmp <> 0 Then
                Load tile(tmp)
                tile(tmp).Left = (col * 480)
                tile(tmp).Top = (row * 480)
            End If
        Next col
    Next row
    loadMap ("storex3y7.txt")
    Call resetMapPos
    Call makeMap
    man.Left = tile((charY * 20) + charX).Left
    man.Top = tile((charY * 20) + charX).Top
    Call showMap
End Sub



Function makeMap()
    For row = 0 To 19
        For col = 0 To 19
            tile((row * 20) + col).Picture = landtile(currentMap(col, row)).Picture
            If currentMap(col, row) >= 120 And currentMap(col, row) <= 125 Then
                landtile(currentMap(col, row)).ZOrder (1)
            End If
        Next col
    Next row
End Function

Function showMap()
    For a = 0 To 399
        tile(a).Visible = True
    Next a
    For a = 0 To 19
        For b = 0 To 19
            If currentMap(b, a) = 114 And treasure(mapX, mapY, b, a) = 0 Then
                currentMap(b, a) = 115
                Call makeMap
            End If
        Next b
    Next a
    man.Visible = True
End Function

Function hideMap()
    man.Visible = False
    For a = 0 To 399
        tile(a).Visible = False
    Next a
End Function


Function resetMapPos()
Dim tmp As Integer
    For row = 0 To 19
        For col = 0 To 19
            tmp = (row * 20) + col
            If charY < 16 And charY > 3 Then
                tile(tmp).Top = (row * 480) - (480 * (charY - 4))
            End If
            If charY > 15 Then
                tile(tmp).Top = (row * 480) - (480 * 11)
            End If
            If charY < 4 Then
                tile(tmp).Top = (row * 480)
            End If
            If charX < 16 And charX > 3 Then
                tile(tmp).Left = (col * 480) - (480 * (charX - 4))
            End If
            If charX > 15 Then
                tile(tmp).Left = (col * 480) - (480 * 11)
            End If
            If charX < 4 Then
                tile(tmp).Left = (col * 480)
            End If

        Next col
    Next row
End Function


'***************************************************************************
'* this reads the array and returns whether or not the character can
'* walk to the next space, or if something is blocking them
'***************************************************************************



Function canWalk(x As Integer, y As Integer) As Boolean

    walkctr = (walkctr + 1) Mod 2
    canWalk = True
    If currentMap(x, y) = WATER Then
        canWalk = False
    End If
    If currentMap(x, y) = MOUNTAIN Or currentMap(x, y) = 66 Then
        canWalk = False
    End If
    If currentMap(x, y) = ROOF Then
        canWalk = False
    End If
    If currentMap(x, y) = TREEBLOCK Then
        canWalk = False
    End If
    If currentMap(x, y) = WALL Then
        canWalk = False
    End If
    If currentMap(x, y) >= PERSON0 And currentMap(x, y) < 31 Then
        canWalk = False
    End If
    
    'roof
    If currentMap(x, y) >= 68 And currentMap(x, y) <= 85 Then
        canWalk = False
    End If

    If currentMap(x, y) >= 160 And currentMap(x, y) <= 163 Then
        canWalk = False
    End If
    If currentMap(x, y) >= 125 And currentMap(x, y) <= 151 Then
        canWalk = False
    End If
    If currentMap(x, y) >= 104 And currentMap(x, y) <= 119 Then
        canWalk = False
    End If
    If currentMap(x, y) >= 91 And currentMap(x, y) <= 102 Then
        canWalk = False
    End If
    If currentMap(x, y) >= 173 And currentMap(x, y) <= 174 Then
        canWalk = False
    End If
    If currentMap(x, y) >= 176 And currentMap(x, y) <= 178 Then
        canWalk = False
    End If
    If currentMap(x, y) >= 182 And currentMap(x, y) <= 184 Then
        canWalk = False
    End If
    If currentMap(x, y) = 187 Or currentMap(x, y) = 171 Then
        canWalk = False
    End If
    If currentMap(x, y) >= 189 And currentMap(x, y) <= 190 Then
        canWalk = False
    End If
    
    'people
    If currentMap(x, y) >= 199 And currentMap(x, y) <= 230 Then
        canWalk = False
    End If
    
    
    'castle stuff
    If currentMap(x, y) >= 241 And currentMap(x, y) <= 243 Then
        canWalk = False
    End If
    If currentMap(x, y) >= 245 And currentMap(x, y) <= 247 Then
        canWalk = False
    End If
    
End Function


Private Sub movement_Timer()
'**********************************************
'This makes the people in town move so they don't
' look like they're dead
'**********************************************
    For a = 0 To 19
        For b = 0 To 19
            If (currentMap(b, a) > 13 And currentMap(b, a) < 30) Or (currentMap(b, a) >= 199 And currentMap(b, a) <= 230) Then
                Dim tmp As Integer
                tmp = currentMap(b, a)
                tile((a * 20) + b).Picture = landtile(tmp + cntr).Picture
            End If
        Next b
    Next a
    If cntr = 0 Then
        cntr = 1
    Else
        cntr = 0
    End If

End Sub

Private Sub quitg_Click()
    End
End Sub

Private Sub RPG_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then
        Load gameMenu
        gameMenu.Width = Me.Width
        gameMenu.Height = Me.Height
        gameMenu.Top = Me.Top
        gameMenu.Left = Me.Left
        gameMenu.Show
        Me.Enabled = False

    End If


    Call hideMap
    Label7.Caption = charY
    Label2.Caption = charX
    Dim b As Integer
    Dim c As Integer
    Dim d As Integer
    Dim e As Integer
    Dim tmpx As Integer
    Dim tmpy As Integer
    
    If KeyCode = vbKeySpace Then
        If direction = NORTH Then
            tmpx = charX
            tmpy = charY - 1
        End If
        If direction = EAST Then
            tmpx = charX + 1
            tmpy = charY
        End If
        If direction = SOUTH Then
            tmpx = charX
            tmpy = charY + 1
        End If
        If direction = WEST Then
            tmpx = charX - 1
            tmpy = charY
        End If
        dialogue.Caption = ""
        
        'load a store if possible
        If storeIDs(mapX, mapY, tmpx, tmpy) <> 0 Then
            getStore (storeIDs(mapX, mapY, tmpx, tmpy))
            Load shop
            shop.Width = Me.Width
            shop.Height = Me.Height
            shop.Top = Me.Top
            shop.Left = Me.Left
            shop.Show
            Me.Enabled = False
        End If
            

        If (currentMap(tmpx, tmpy) > 13 And currentMap(tmpx, tmpy) < 30) Or (currentMap(tmpx, tmpy) >= 199 And currentMap(tmpx, tmpy) <= 230) Then
            eraseTalk = 0
            b = affects(mapX, mapY, tmpx, tmpy, 0)
            c = affects(mapX, mapY, tmpx, tmpy, 1)
            d = affects(mapX, mapY, tmpx, tmpy, 2)
            e = affects(mapX, mapY, tmpx, tmpy, 3)
            Dim doesAff As Boolean
            doesAff = False
            'see if dialogue affects anyone else
            For a = 0 To 3
                If affects(mapX, mapY, tmpx, tmpy, a) <> 0 Then
                    doesAff = True
                    a = 4
                End If
            Next a
                            
            If doesAff = True Then
                For a = 0 To 19
                    progressSpeech(b, c, d, e, a) = progressSpeech(b, c, d, e, a + 1)
                    speech(b, c, d, e, a) = speech(b, c, d, e, a + 1)
                Next a
                
                
                
                For a = 0 To 20
                   If progressSpeech(b, c, d, e, a) = 0 Then
                      progressSpeech(b, c, d, e, a) = 1
                       a = 21
                    End If
                Next a
            End If
            If progressSpeech(mapX, mapY, tmpx, tmpy, 0) = 1 Or doesAff = True Then
                For a = 0 To 3
                    For f = 0 To 399
                        affects(mapX, mapY, tmpx, tmpy, f) = affects(mapX, mapY, tmpx, tmpy, f + 1)
                    Next f
                Next a
            End If

            If speech(mapX, mapY, tmpx, tmpy, 0) = "" Then
                dialogue.Caption = "I have nothing to say to you."
            Else
                dialogue.Caption = speech(mapX, mapY, tmpx, tmpy, 0)
            End If
            isTalking = False
            If progressSpeech(mapX, mapY, tmpx, tmpy, 0) = 1 Then
                isTalking = True
            End If
            'see if talking caused an event
            If doEvent(mapX, mapY, tmpx, tmpy, 0) <> 0 Then
                Call perform(doEvent(mapX, mapY, tmpx, tmpy, 0))
                doEvent(mapX, mapY, tmpx, tmpy, 0) = 0
            End If
            If progressSpeech(mapX, mapY, tmpx, tmpy, 0) = 1 Then
                For a = 0 To 19
                    speech(mapX, mapY, tmpx, tmpy, a) = speech(mapX, mapY, tmpx, tmpy, a + 1)
                    progressSpeech(mapX, mapY, tmpx, tmpy, a) = progressSpeech(mapX, mapY, tmpx, tmpy, a + 1)
                    doEvent(mapX, mapY, tmpx, tmpy, a) = doEvent(mapX, mapY, tmpx, tmpy, a + 1)
                Next a
            End If
            

            
        End If
    'for treasure!!!!
        If currentMap(tmpx, tmpy) = 114 Then
             currentMap(tmpx, tmpy) = 115
             Call makeMap
             Call getTreasure(treasure(mapX, mapY, tmpx, tmpy))
             treasure(mapX, mapY, tmpx, tmpy) = 0
        Else
            If currentMap(tmpx, tmpy) = 115 Then
                dialogue.Caption = "Nothing in here."
            End If
        End If
    End If
    
    
    
    
    'delay for time dialogue is erased from label
    If KeyCode >= 37 And KeyCode <= 40 And isTalking = False Then
        eraseTalk = eraseTalk + 1
        If eraseTalk = 6 Then
            dialogue.Caption = ""
            eraseTalk = 0
        End If
    End If
    
    If isTalking = False Then
        If KeyCode = vbKeyRight Then
            direction = EAST
            If charX < 19 Then
                If canWalk(charX + 1, charY) Then
                    If tile(11).Left > 0 And charX > 3 Then
                        For a = 0 To 399
                            tile(a).Left = tile(a).Left - 480
                        Next a
                    End If
                    charX = charX + 1
                End If
            Else
                Call hideMap
                mapX = (mapX + 1) Mod 10
                Call loadMap(mapName(mapX, mapY))
                Call makeMap
                charX = 0
                Call resetMapPos
                        man.Left = tile((charY * 20) + charX).Left
                        man.Top = tile((charY * 20) + charX).Top
                Call showMap
            End If
        End If
        If KeyCode = vbKeyLeft Then
        direction = WEST
            If charX > 0 Then
                If canWalk(charX - 1, charY) Then
                    If tile(0).Left < 0 And charX < 16 Then
                        For a = 0 To 399
                            tile(a).Left = tile(a).Left + 480
                        Next a
                    End If
                    charX = charX - 1
                End If
            Else
                Call hideMap
                mapX = (mapX - 1) Mod 10
                Call loadMap(mapName(mapX, mapY))
                Call makeMap
                charX = 19
                Call resetMapPos
                man.Left = tile((charY * 20) + charX).Left
                man.Top = tile((charY * 20) + charX).Top
                Call showMap
            End If
        End If
        If KeyCode = vbKeyUp Then
        direction = NORTH
             If charY > 0 Then
                 If canWalk(charX, charY - 1) Then
                     If tile(0).Top < 0 And charY < 16 Then
                         For a = 0 To 399
                             tile(a).Top = tile(a).Top + 480
                         Next a
                     End If
                     charY = charY - 1
                 End If
             Else
                 Call hideMap
                 mapY = (mapY - 1) Mod 10
                 Call loadMap(mapName(mapX, mapY))
                 Call makeMap
                 charY = 19
                 Call resetMapPos
                 man.Left = tile((charY * 20) + charX).Left
                 man.Top = tile((charY * 20) + charX).Top
                 Call showMap
            End If
       End If
        If KeyCode = vbKeyDown Then
            direction = SOUTH
            If charY < 19 Then
                If canWalk(charX, charY + 1) Then
                    If tile(220).Top > 0 And charY > 3 Then
                        For a = 0 To 399
                            tile(a).Top = tile(a).Top - 480
                        Next a
                    End If
                    charY = charY + 1
                End If
            Else
                If talked = True Then
                dialogue.Caption = ""
                If inHotSpot = False Then
                    Call hideMap
                    mapY = (mapY + 1) Mod 10
                    Call loadMap(mapName(mapX, mapY))
                    Call makeMap
                    charY = 0
                    Call resetMapPos
                    man.Left = tile((charY * 20) + charX).Left
                    man.Top = tile((charY * 20) + charX).Top
                    Call showMap
                Else
                    If inDoor = False Then
                        Call hideMap
                        inHotSpot = False
                        movement.Enabled = False
                        Call loadMap(mapName(mapX, mapY))
                        Call makeMap
                        charY = lasty
                        charX = lastx
                        Call resetMapPos
                        man.Left = tile((charY * 20) + charX).Left
                        man.Top = tile((charY * 20) + charX).Top
                        Call showMap
                    Else
                        Call hideMap
                        inDoor = False
                        Call loadMap(mapHotSpot(mapX, mapY, lastx, lasty + 1))
                        Call makeMap
                        charY = lastdoory
                        charX = lastdoorx
                        Call resetMapPos
                        man.Left = tile((charY * 20) + charX).Left
                        man.Top = tile((charY * 20) + charX).Top
                        Call showMap
                    End If
                End If
                End If
            End If
        End If
    End If
    Call makeMap
    man.Left = tile((charY * 20) + charX).Left
    man.Top = tile((charY * 20) + charX).Top
    Call showMap


    'go into a city
    If (mapHotSpot(mapX, mapY, charX, charY) <> "" Or (currentMap(charX, charY) >= 248 And currentMap(charX, charY) <= 251)) And inHotSpot = False And inDoor = False Then
        lastx = charX
        lasty = charY - 1
        inHotSpot = True
        movement = True
        Call hideMap
        Call loadMap(mapHotSpot(mapX, mapY, charX, charY))
        Call makeMap
        charY = 19
        charX = 18
        Call resetMapPos
        man.Left = tile((charY * 20) + charX).Left
        man.Top = tile((charY * 20) + charX).Top
        Call showMap
    End If
       
'-----CHECK IF DOOR------
    If inHotSpot = True Then
        If currentMap(charX, charY) = DOOR Then
            inDoor = True
            lastdoorx = charX
            lastdoory = charY + 1
            Call hideMap
            Call loadMap(hotDoor(mapX, mapY, charX, charY))
            Call makeMap
            If keepPos(mapX, mapY, charX, charY) = 0 Then
                charY = 19
                charX = 15
            Else
                charY = 19
            End If
            Call resetMapPos
            man.Left = tile((charY * 20) + charX).Left
            man.Top = tile((charY * 20) + charX).Top
            Call showMap
        End If
    End If
'-----CHECK IF STAIR that relocated character in stairs------
    If currentMap(charX, charY) = STAIR Then
    If hotStair(mapX, mapY, charX, charY, 0) <> 0 And hotStair(mapX, mapY, charX, charY, 1) <> 0 Then
        Dim tx As Integer
        Dim ty As Integer
                tx = hotStair(mapX, mapY, charX, charY, 0)
                ty = hotStair(mapX, mapY, charX, charY, 1)
                charX = tx
                charY = ty
                Call hideMap
                Call resetMapPos
                man.Left = tile((charY * 20) + charX).Left
                man.Top = tile((charY * 20) + charX).Top
                Call showMap
        End If
    End If
    
'---check if it a stair case the needs to load a new map
    If talked = True And loadstair(mapX, mapY, charX, charY) <> "" And currentMap(charX, charY) = STAIR Then
        Call hideMap
        Call loadMap(loadstair(mapX, mapY, charX, charY))
        Call makeMap
        For a = 0 To 19
            For b = 0 To 19
                If currentMap(b, a) = STAIR Then
                    charY = a
                    charX = b
                End If
            Next b
        Next a
        Call resetMapPos
        man.Left = tile((charY * 20) + charX).Left
        man.Top = tile((charY * 20) + charX).Top
        Call showMap
    End If

    If currentMap(charX, charY) >= 120 And currentMap(charX, charY) <= 124 Then
        man.Visible = False
    Else
        man.Visible = True
        If currentMap(charX, charY) >= 180 And currentMap(charX, charY) <= 181 Then
            man.Visible = False
        End If
        If currentMap(charX, charY) = 103 Or currentMap(charX, charY) = 244 Then
            man.ZOrder (1)
        Else
            man.ZOrder (0)
        End If
    End If
    
    
    'random battles!!!
    If KeyCode >= 37 And KeyCode <= 40 And inHotSpot = False And inDoor = False Then
        Randomize
        Dim tmp As Integer
        tmp = Int(Rnd * 25)
        If tmp = 0 Then
            Me.Enabled = False
            Load Battle
            Battle.Left = Me.Left
            Battle.Top = Me.Top
            Battle.Width = Me.Width
            Battle.Height = Me.Height
            Battle.Show
        End If
    End If
    man.Picture = people.guy(direction + (4 * walkctr)).Picture
End Sub
