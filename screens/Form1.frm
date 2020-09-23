VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9885
   ClientLeft      =   930
   ClientTop       =   1155
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   14325
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox mname2 
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   7320
      Width           =   975
   End
   Begin VB.TextBox mname 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6840
      Width           =   975
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   215
      Left            =   13200
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   216
      Left            =   13680
      Picture         =   "Form1.frx":03B6
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   217
      Left            =   13200
      Picture         =   "Form1.frx":076C
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   218
      Left            =   13680
      Picture         =   "Form1.frx":0B22
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   219
      Left            =   13200
      Picture         =   "Form1.frx":0ED8
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   220
      Left            =   13680
      Picture         =   "Form1.frx":128E
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   221
      Left            =   13200
      Picture         =   "Form1.frx":1644
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   222
      Left            =   13680
      Picture         =   "Form1.frx":1986
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   223
      Left            =   13200
      Picture         =   "Form1.frx":1CC8
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   224
      Left            =   13680
      Picture         =   "Form1.frx":200A
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   225
      Left            =   13200
      Picture         =   "Form1.frx":234C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   226
      Left            =   13680
      Picture         =   "Form1.frx":2702
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   227
      Left            =   13200
      Picture         =   "Form1.frx":2AB8
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   228
      Left            =   13680
      Picture         =   "Form1.frx":2E6E
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   229
      Left            =   13200
      Picture         =   "Form1.frx":3224
      Stretch         =   -1  'True
      Top             =   480
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   230
      Left            =   13680
      Picture         =   "Form1.frx":35E0
      Stretch         =   -1  'True
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "map name"
      Height          =   255
      Left            =   6360
      TabIndex        =   9
      Top             =   6480
      Width           =   975
   End
   Begin VB.Image tile2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   251
      Left            =   10800
      Picture         =   "Form1.frx":399B
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   250
      Left            =   10440
      Picture         =   "Form1.frx":3CDD
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   249
      Left            =   10800
      Picture         =   "Form1.frx":401F
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   248
      Left            =   10440
      Picture         =   "Form1.frx":4361
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   247
      Left            =   9600
      Picture         =   "Form1.frx":46A3
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   246
      Left            =   9360
      Picture         =   "Form1.frx":49E5
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   245
      Left            =   9840
      Picture         =   "Form1.frx":4D27
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   244
      Left            =   8880
      Picture         =   "Form1.frx":5069
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   243
      Left            =   8880
      Picture         =   "Form1.frx":53AB
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   242
      Left            =   8880
      Picture         =   "Form1.frx":56ED
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   241
      Left            =   8880
      Picture         =   "Form1.frx":5A2F
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   240
      Left            =   8400
      Picture         =   "Form1.frx":5D71
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   239
      Left            =   7320
      Picture         =   "Form1.frx":60B3
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   238
      Left            =   7680
      Picture         =   "Form1.frx":63F5
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   237
      Left            =   7320
      Picture         =   "Form1.frx":6737
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   236
      Left            =   7320
      Picture         =   "Form1.frx":6A79
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   235
      Left            =   7680
      Picture         =   "Form1.frx":6DBB
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   234
      Left            =   7680
      Picture         =   "Form1.frx":70FD
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   233
      Left            =   8040
      Picture         =   "Form1.frx":747F
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   232
      Left            =   8040
      Picture         =   "Form1.frx":77C1
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   231
      Left            =   8040
      Picture         =   "Form1.frx":7B03
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   214
      Left            =   12720
      Picture         =   "Form1.frx":7E45
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   213
      Left            =   12360
      Picture         =   "Form1.frx":81FB
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   212
      Left            =   12720
      Picture         =   "Form1.frx":85B1
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   211
      Left            =   12360
      Picture         =   "Form1.frx":88F3
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   210
      Left            =   12720
      Picture         =   "Form1.frx":8C35
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   209
      Left            =   12360
      Picture         =   "Form1.frx":8F77
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   208
      Left            =   12720
      Picture         =   "Form1.frx":92B9
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   207
      Left            =   12360
      Picture         =   "Form1.frx":95FB
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   206
      Left            =   12720
      Picture         =   "Form1.frx":993D
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   205
      Left            =   12360
      Picture         =   "Form1.frx":9C7F
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   204
      Left            =   12720
      Picture         =   "Form1.frx":9FC1
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   203
      Left            =   12360
      Picture         =   "Form1.frx":A377
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   202
      Left            =   12720
      Picture         =   "Form1.frx":A72D
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   201
      Left            =   12360
      Picture         =   "Form1.frx":AAE3
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   200
      Left            =   12720
      Picture         =   "Form1.frx":AE99
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   199
      Left            =   12360
      Picture         =   "Form1.frx":B1DB
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   14
      Left            =   11400
      Picture         =   "Form1.frx":B51D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   15
      Left            =   11760
      Picture         =   "Form1.frx":B85F
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   16
      Left            =   11400
      Picture         =   "Form1.frx":BBA1
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   17
      Left            =   11760
      Picture         =   "Form1.frx":BF57
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   18
      Left            =   11400
      Picture         =   "Form1.frx":C30D
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   19
      Left            =   11760
      Picture         =   "Form1.frx":C64F
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   20
      Left            =   11400
      Picture         =   "Form1.frx":C991
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   21
      Left            =   11760
      Picture         =   "Form1.frx":CCD3
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   22
      Left            =   11400
      Picture         =   "Form1.frx":D015
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   23
      Left            =   11760
      Picture         =   "Form1.frx":D357
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   24
      Left            =   11400
      Picture         =   "Form1.frx":D699
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   25
      Left            =   11760
      Picture         =   "Form1.frx":DA4F
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   26
      Left            =   11400
      Picture         =   "Form1.frx":DE05
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   27
      Left            =   11760
      Picture         =   "Form1.frx":E1BB
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   28
      Left            =   11400
      Picture         =   "Form1.frx":E571
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   29
      Left            =   11760
      Picture         =   "Form1.frx":E927
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   198
      Left            =   11160
      Picture         =   "Form1.frx":ECDD
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   197
      Left            =   10800
      Picture         =   "Form1.frx":F01F
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   196
      Left            =   11160
      Picture         =   "Form1.frx":F361
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   195
      Left            =   10800
      Picture         =   "Form1.frx":F6A3
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   194
      Left            =   10440
      Picture         =   "Form1.frx":F9E5
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   193
      Left            =   10080
      Picture         =   "Form1.frx":FD27
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   192
      Left            =   10440
      Picture         =   "Form1.frx":10069
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   191
      Left            =   10080
      Picture         =   "Form1.frx":103AB
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   40
      Left            =   1680
      Picture         =   "Form1.frx":106ED
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   66
      Left            =   10680
      Picture         =   "Form1.frx":10AA3
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   67
      Left            =   10680
      Picture         =   "Form1.frx":10DE5
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   68
      Left            =   11040
      Picture         =   "Form1.frx":11127
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   69
      Left            =   10680
      Picture         =   "Form1.frx":11469
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   70
      Left            =   10320
      Picture         =   "Form1.frx":117AB
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   71
      Left            =   10320
      Picture         =   "Form1.frx":11AED
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   72
      Left            =   11040
      Picture         =   "Form1.frx":11E2F
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   73
      Left            =   10320
      Picture         =   "Form1.frx":12171
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   74
      Left            =   11040
      Picture         =   "Form1.frx":124B3
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   190
      Left            =   14400
      Picture         =   "Form1.frx":127F5
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   189
      Left            =   12120
      Picture         =   "Form1.frx":12B37
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   188
      Left            =   12120
      Picture         =   "Form1.frx":12E79
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   187
      Left            =   13680
      Picture         =   "Form1.frx":131BB
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   186
      Left            =   14400
      Picture         =   "Form1.frx":134FD
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   185
      Left            =   13680
      Picture         =   "Form1.frx":1383F
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   184
      Left            =   13320
      Picture         =   "Form1.frx":13B81
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   183
      Left            =   14760
      Picture         =   "Form1.frx":13EC3
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   182
      Left            =   12840
      Picture         =   "Form1.frx":14205
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   181
      Left            =   12480
      Picture         =   "Form1.frx":14547
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   180
      Left            =   14040
      Picture         =   "Form1.frx":14889
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   179
      Left            =   12840
      Picture         =   "Form1.frx":14BCB
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   178
      Left            =   13320
      Picture         =   "Form1.frx":14F0D
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   177
      Left            =   14760
      Picture         =   "Form1.frx":1524F
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   176
      Left            =   13680
      Picture         =   "Form1.frx":15591
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   175
      Left            =   13680
      Picture         =   "Form1.frx":158D3
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   174
      Left            =   14400
      Picture         =   "Form1.frx":15C15
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   173
      Left            =   14040
      Picture         =   "Form1.frx":15F57
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   172
      Left            =   14400
      Picture         =   "Form1.frx":162D9
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   171
      Left            =   14040
      Picture         =   "Form1.frx":1661B
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   170
      Left            =   2520
      Picture         =   "Form1.frx":1695D
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   169
      Left            =   2520
      Picture         =   "Form1.frx":16C9F
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   168
      Left            =   2880
      Picture         =   "Form1.frx":16FE1
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   167
      Left            =   2520
      Picture         =   "Form1.frx":17323
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   166
      Left            =   2880
      Picture         =   "Form1.frx":17665
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   165
      Left            =   2880
      Picture         =   "Form1.frx":179A7
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   164
      Left            =   3240
      Picture         =   "Form1.frx":17CE9
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   163
      Left            =   3240
      Picture         =   "Form1.frx":1802B
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   162
      Left            =   3240
      Picture         =   "Form1.frx":1836D
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   161
      Left            =   3240
      Picture         =   "Form1.frx":186AF
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   160
      Left            =   3240
      Picture         =   "Form1.frx":189F1
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   159
      Left            =   2160
      Picture         =   "Form1.frx":18D33
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   158
      Left            =   2160
      Picture         =   "Form1.frx":19075
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   157
      Left            =   720
      Picture         =   "Form1.frx":193B7
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   156
      Left            =   720
      Picture         =   "Form1.frx":196F9
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   155
      Left            =   1680
      Picture         =   "Form1.frx":19A3B
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   154
      Left            =   1680
      Picture         =   "Form1.frx":19D7D
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   153
      Left            =   1680
      Picture         =   "Form1.frx":1A0BF
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   152
      Left            =   1680
      Picture         =   "Form1.frx":1A401
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   151
      Left            =   1080
      Picture         =   "Form1.frx":1A743
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   150
      Left            =   360
      Picture         =   "Form1.frx":1AAF9
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   149
      Left            =   360
      Picture         =   "Form1.frx":1AEAF
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   148
      Left            =   1080
      Picture         =   "Form1.frx":1B265
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   147
      Left            =   360
      Picture         =   "Form1.frx":1B61B
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   146
      Left            =   1080
      Picture         =   "Form1.frx":1B9D1
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   145
      Left            =   360
      Picture         =   "Form1.frx":1BD87
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   144
      Left            =   3840
      Picture         =   "Form1.frx":1C13D
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   143
      Left            =   4200
      Picture         =   "Form1.frx":1C4F3
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   142
      Left            =   3840
      Picture         =   "Form1.frx":1C8A9
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   141
      Left            =   3840
      Picture         =   "Form1.frx":1CC5F
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   140
      Left            =   4200
      Picture         =   "Form1.frx":1D015
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   139
      Left            =   4560
      Picture         =   "Form1.frx":1D3CB
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   138
      Left            =   4560
      Picture         =   "Form1.frx":1D781
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   137
      Left            =   4560
      Picture         =   "Form1.frx":1DB37
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   136
      Left            =   5040
      Picture         =   "Form1.frx":1DEED
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   135
      Left            =   5040
      Picture         =   "Form1.frx":1E2A3
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   134
      Left            =   6120
      Picture         =   "Form1.frx":1E659
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   133
      Left            =   6120
      Picture         =   "Form1.frx":1EA0F
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   132
      Left            =   6120
      Picture         =   "Form1.frx":1EDC5
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   131
      Left            =   5760
      Picture         =   "Form1.frx":1F17B
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   130
      Left            =   5760
      Picture         =   "Form1.frx":1F531
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   129
      Left            =   5400
      Picture         =   "Form1.frx":1F8E7
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   128
      Left            =   5400
      Picture         =   "Form1.frx":1FC9D
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   127
      Left            =   5760
      Picture         =   "Form1.frx":20053
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   126
      Left            =   5400
      Picture         =   "Form1.frx":20409
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   125
      Left            =   2520
      Picture         =   "Form1.frx":207BF
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   124
      Left            =   2880
      Picture         =   "Form1.frx":20B01
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   123
      Left            =   2520
      Picture         =   "Form1.frx":20E43
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   122
      Left            =   2880
      Picture         =   "Form1.frx":21185
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   121
      Left            =   2160
      Picture         =   "Form1.frx":214C7
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   120
      Left            =   2160
      Picture         =   "Form1.frx":21809
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   119
      Left            =   5760
      Picture         =   "Form1.frx":21B4B
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   118
      Left            =   5760
      Picture         =   "Form1.frx":21E8D
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   117
      Left            =   6480
      Picture         =   "Form1.frx":221CF
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   116
      Left            =   4680
      Picture         =   "Form1.frx":22511
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   115
      Left            =   6840
      Picture         =   "Form1.frx":22853
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   114
      Left            =   6480
      Picture         =   "Form1.frx":22B95
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   113
      Left            =   1800
      Picture         =   "Form1.frx":22ED7
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   112
      Left            =   1440
      Picture         =   "Form1.frx":23219
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   111
      Left            =   1440
      Picture         =   "Form1.frx":2355B
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   110
      Left            =   1800
      Picture         =   "Form1.frx":2389D
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   109
      Left            =   6480
      Picture         =   "Form1.frx":23BDF
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   108
      Left            =   6840
      Picture         =   "Form1.frx":23F95
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   107
      Left            =   6480
      Picture         =   "Form1.frx":2434B
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   106
      Left            =   6480
      Picture         =   "Form1.frx":24701
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   105
      Left            =   6840
      Picture         =   "Form1.frx":24AB7
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   104
      Left            =   6840
      Picture         =   "Form1.frx":24E6D
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   103
      Left            =   5400
      Picture         =   "Form1.frx":25223
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   102
      Left            =   5400
      Picture         =   "Form1.frx":25925
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   101
      Left            =   1440
      Picture         =   "Form1.frx":26027
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   100
      Left            =   1440
      Picture         =   "Form1.frx":263DD
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   99
      Left            =   1080
      Picture         =   "Form1.frx":26793
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   98
      Left            =   5040
      Picture         =   "Form1.frx":26B49
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   97
      Left            =   5040
      Picture         =   "Form1.frx":26EFF
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   96
      Left            =   720
      Picture         =   "Form1.frx":272B5
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   95
      Left            =   720
      Picture         =   "Form1.frx":2766B
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   94
      Left            =   360
      Picture         =   "Form1.frx":27A21
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   93
      Left            =   360
      Picture         =   "Form1.frx":27DD7
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   92
      Left            =   360
      Picture         =   "Form1.frx":2818D
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   91
      Left            =   360
      Picture         =   "Form1.frx":28543
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   90
      Left            =   3120
      Picture         =   "Form1.frx":288F9
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   89
      Left            =   3480
      Picture         =   "Form1.frx":2A37B
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "countertop"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   88
      Left            =   2040
      Picture         =   "Form1.frx":2BDFD
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "counter"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   87
      Left            =   3840
      Picture         =   "Form1.frx":2D87F
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   86
      Left            =   3360
      Picture         =   "Form1.frx":2F301
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   85
      Left            =   3240
      Picture         =   "Form1.frx":30BB3
      Stretch         =   -1  'True
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   84
      Left            =   3240
      Picture         =   "Form1.frx":32465
      Stretch         =   -1  'True
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   83
      Left            =   2880
      Picture         =   "Form1.frx":33D17
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   82
      Left            =   2880
      Picture         =   "Form1.frx":355C9
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   81
      Left            =   2880
      Picture         =   "Form1.frx":36E7B
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   80
      Left            =   2520
      Picture         =   "Form1.frx":3872D
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   79
      Left            =   3240
      Picture         =   "Form1.frx":3A097
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   78
      Left            =   2520
      Picture         =   "Form1.frx":3BA01
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   77
      Left            =   3240
      Picture         =   "Form1.frx":3D36B
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   76
      Left            =   3240
      Picture         =   "Form1.frx":3ECD5
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   75
      Left            =   2520
      Picture         =   "Form1.frx":4063F
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   65
      Left            =   11040
      Picture         =   "Form1.frx":41FA9
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   64
      Left            =   10320
      Picture         =   "Form1.frx":422EB
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   63
      Left            =   10320
      Picture         =   "Form1.frx":4262D
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   62
      Left            =   11040
      Picture         =   "Form1.frx":4296F
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   61
      Left            =   11040
      Picture         =   "Form1.frx":42CB1
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   60
      Left            =   10680
      Picture         =   "Form1.frx":42FF3
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   59
      Left            =   10320
      Picture         =   "Form1.frx":43335
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   58
      Left            =   10680
      Picture         =   "Form1.frx":43677
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   57
      Left            =   3600
      Picture         =   "Form1.frx":439B9
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   56
      Left            =   11880
      Picture         =   "Form1.frx":43CFB
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   55
      Left            =   12240
      Picture         =   "Form1.frx":4403D
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   54
      Left            =   11880
      Picture         =   "Form1.frx":4437F
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   53
      Left            =   11520
      Picture         =   "Form1.frx":446C1
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   52
      Left            =   11520
      Picture         =   "Form1.frx":44A03
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   51
      Left            =   12240
      Picture         =   "Form1.frx":44D45
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   50
      Left            =   11520
      Picture         =   "Form1.frx":45087
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   49
      Left            =   12240
      Picture         =   "Form1.frx":453C9
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   48
      Left            =   12240
      Picture         =   "Form1.frx":4570B
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   47
      Left            =   11520
      Picture         =   "Form1.frx":45A4D
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   46
      Left            =   12240
      Picture         =   "Form1.frx":45D8F
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   45
      Left            =   11520
      Picture         =   "Form1.frx":460D1
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   44
      Left            =   11520
      Picture         =   "Form1.frx":46413
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   43
      Left            =   11880
      Picture         =   "Form1.frx":46755
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   42
      Left            =   12240
      Picture         =   "Form1.frx":46A97
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   41
      Left            =   11880
      Picture         =   "Form1.frx":46DD9
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   39
      Left            =   10680
      Picture         =   "Form1.frx":4711B
      Stretch         =   -1  'True
      Top             =   8760
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   38
      Left            =   11040
      Picture         =   "Form1.frx":4745D
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   37
      Left            =   10680
      Picture         =   "Form1.frx":4779F
      Stretch         =   -1  'True
      Top             =   9480
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   36
      Left            =   10320
      Picture         =   "Form1.frx":47AE1
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   35
      Left            =   10320
      Picture         =   "Form1.frx":47E23
      Stretch         =   -1  'True
      Top             =   9480
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   34
      Left            =   11040
      Picture         =   "Form1.frx":48165
      Stretch         =   -1  'True
      Top             =   9480
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   33
      Left            =   10320
      Picture         =   "Form1.frx":484A7
      Stretch         =   -1  'True
      Top             =   8760
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   32
      Left            =   11040
      Picture         =   "Form1.frx":487E9
      Stretch         =   -1  'True
      Top             =   8760
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   31
      Left            =   5400
      Picture         =   "Form1.frx":48B2B
      Stretch         =   -1  'True
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "current"
      Height          =   255
      Left            =   11640
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.Image current 
      Height          =   375
      Left            =   11640
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   0
      Left            =   11520
      Picture         =   "Form1.frx":4A5AD
      Stretch         =   -1  'True
      Top             =   8760
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   1
      Left            =   10680
      Picture         =   "Form1.frx":4A8EF
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   2
      Left            =   11880
      Picture         =   "Form1.frx":4AC31
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   3
      Left            =   11880
      Picture         =   "Form1.frx":4AF73
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   4
      Left            =   3840
      Picture         =   "Form1.frx":4B2B5
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   5
      Left            =   5640
      Picture         =   "Form1.frx":4CD37
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image character 
      Height          =   375
      Left            =   9720
      Picture         =   "Form1.frx":4E7B9
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   10
      Left            =   4080
      Picture         =   "Form1.frx":5023B
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   "block"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   9
      Left            =   3120
      Picture         =   "Form1.frx":5057D
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   8
      Left            =   1080
      Picture         =   "Form1.frx":51FFF
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   7
      Left            =   2520
      Picture         =   "Form1.frx":523B5
      Stretch         =   -1  'True
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   6
      Left            =   4080
      Picture         =   "Form1.frx":53E37
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   12
      Left            =   1080
      Picture         =   "Form1.frx":558B9
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   11
      Left            =   720
      Picture         =   "Form1.frx":55C6F
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   13
      Left            =   12480
      Picture         =   "Form1.frx":56025
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image landtile 
      Height          =   375
      Index           =   30
      Left            =   4800
      Picture         =   "Form1.frx":56367
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "map name"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   6600
      Width           =   975
   End
   Begin VB.Image tile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":57DE9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu saveMap 
         Caption         =   "Save1"
      End
      Begin VB.Menu saveMap2 
         Caption         =   "Save2"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public layval As Integer


Private Sub Command1_Click()
    If mname.Text <> "" Then
        Call loadThis(mname)
    End If
    For row = 0 To 19
        For col = 0 To 19
            tile((row * 20) + col).Picture = landtile(currentMap(col, row)).Picture
        Next col
    Next row
End Sub


Private Sub Command2_Click()
    If mname2.Text <> "" Then
        Call loadThis2(mname2)
    End If
    For row = 0 To 19
        For col = 0 To 19
            tile2((row * 20) + col).Picture = landtile(currentMap2(col, row)).Picture
        Next col
    Next row
End Sub

Private Sub Form_Load()
layval = 0
Dim tmp As Integer
    For row = 0 To 19
        For col = 0 To 19
            tmp = (row * 20) + col
            If tmp <> 0 Then
                Load tile(tmp)
                tile(tmp).Left = col * 240
                tile(tmp).Top = row * 240
                currentMap(col, row) = 2
                Load tile2(tmp)
                tile2(tmp).Left = (col * 240) + 4800
                tile2(tmp).Top = (row * 240)
                currentMap2(col, row) = 2
            End If
        Next col
    Next row
    For a = 0 To 399
        tile(a).Visible = True
        tile2(a).Visible = True
    Next a
End Sub




Private Sub landtile_Click(Index As Integer)
    layval = Index
    current.Picture = landtile(Index).Picture
End Sub


Private Sub saveMap_Click()
    If mname.Text <> "" Then
        Call saveThis(mname)
    End If
End Sub

Private Sub saveMap2_Click()
    If mname2.Text <> "" Then
        Call saveThis2(mname2)
    End If
End Sub

Private Sub tile_Click(Index As Integer)
    Dim row As Integer
    Dim col As Integer
    tile(Index).Picture = current.Picture
    row = Int(Index / 20)
    col = Index Mod 20
    currentMap(col, row) = layval
End Sub

Private Sub tile2_Click(Index As Integer)
    Dim row As Integer
    Dim col As Integer
    tile2(Index).Picture = current.Picture
    row = Int(Index / 20)
    col = Index Mod 20
    currentMap2(col, row) = layval

End Sub
