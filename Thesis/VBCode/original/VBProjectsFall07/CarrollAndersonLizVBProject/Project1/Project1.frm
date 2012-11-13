VERSION 5.00
Begin VB.Form frmMinesweeper 
   Caption         =   "Minesweeper"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   3900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   840
      TabIndex        =   167
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   840
      TabIndex        =   166
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdBackToMathScreen 
      Caption         =   "Back To Math Screen"
      Height          =   495
      Left            =   840
      TabIndex        =   165
      Top             =   5400
      Width           =   1935
   End
   Begin VB.PictureBox picFrown 
      Height          =   495
      Left            =   1560
      Picture         =   "Project1.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   164
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picAI 
      Height          =   375
      Left            =   240
      Picture         =   "Project1.frx":0972
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   163
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAH 
      Height          =   375
      Left            =   240
      Picture         =   "Project1.frx":0F34
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   162
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAG 
      Height          =   375
      Left            =   240
      Picture         =   "Project1.frx":14F6
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   161
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAF 
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   160
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picBI 
      Height          =   375
      Left            =   600
      Picture         =   "Project1.frx":1AB8
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   159
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picBH 
      Height          =   375
      Left            =   600
      Picture         =   "Project1.frx":207A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   158
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picBG 
      Height          =   375
      Left            =   600
      Picture         =   "Project1.frx":263C
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   157
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picBF 
      Height          =   375
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   156
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCI 
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   155
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCH 
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   154
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCG 
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   153
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCF 
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   152
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picDI 
      Height          =   375
      Left            =   1320
      Picture         =   "Project1.frx":2BFE
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   151
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picDH 
      Height          =   375
      Left            =   1320
      Picture         =   "Project1.frx":31C0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   150
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picDG 
      Height          =   375
      Left            =   1320
      Picture         =   "Project1.frx":3782
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   149
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picDF 
      Height          =   375
      Left            =   1320
      Picture         =   "Project1.frx":3D44
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   148
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picEI 
      Height          =   375
      Left            =   1680
      Picture         =   "Project1.frx":4306
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   147
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picEH 
      Height          =   375
      Left            =   1680
      Picture         =   "Project1.frx":48C8
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   146
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picEG 
      Height          =   375
      Left            =   1680
      Picture         =   "Project1.frx":4E8A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   145
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picEF 
      Height          =   375
      Left            =   1680
      Picture         =   "Project1.frx":544C
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   144
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picFI 
      Height          =   375
      Left            =   2040
      Picture         =   "Project1.frx":5A0E
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   143
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picFH 
      Height          =   375
      Left            =   2040
      Picture         =   "Project1.frx":5FD0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   142
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picFG 
      Height          =   375
      Left            =   2040
      Picture         =   "Project1.frx":6592
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   141
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picFF 
      Height          =   375
      Left            =   2040
      Picture         =   "Project1.frx":6B54
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   140
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAE 
      Height          =   375
      Left            =   240
      Picture         =   "Project1.frx":7116
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   139
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picBE 
      Height          =   375
      Left            =   600
      Picture         =   "Project1.frx":76D8
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   138
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCE 
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   137
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picDE 
      Height          =   375
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   136
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picEE 
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   135
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picGI 
      Height          =   375
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   134
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picGH 
      Height          =   375
      Left            =   2400
      Picture         =   "Project1.frx":7C9A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   133
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picGG 
      Height          =   375
      Left            =   2400
      Picture         =   "Project1.frx":825C
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   132
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picGF 
      Height          =   375
      Left            =   2400
      Picture         =   "Project1.frx":881E
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   131
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picFE 
      Height          =   375
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   130
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picGE 
      Height          =   375
      Left            =   2400
      Picture         =   "Project1.frx":8DE0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   129
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picHI 
      Height          =   375
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   128
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picHH 
      Height          =   375
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   127
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picHG 
      Height          =   375
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   126
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picHF 
      Height          =   375
      Left            =   2760
      Picture         =   "Project1.frx":93A2
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   125
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picHE 
      Height          =   375
      Left            =   2760
      Picture         =   "Project1.frx":9964
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   124
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAD 
      Height          =   375
      Left            =   240
      Picture         =   "Project1.frx":9F26
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   123
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picBD 
      Height          =   375
      Left            =   600
      Picture         =   "Project1.frx":A4E8
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   122
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCD 
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   121
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picDD 
      Height          =   375
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   120
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picED 
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   119
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picFD 
      Height          =   375
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   118
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picGD 
      Height          =   375
      Left            =   2400
      Picture         =   "Project1.frx":AAAA
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   117
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picHD 
      Height          =   375
      Left            =   2760
      Picture         =   "Project1.frx":B06C
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   116
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picII 
      Height          =   375
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   115
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picIH 
      Height          =   375
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   114
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picIG 
      Height          =   375
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   113
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picIF 
      Height          =   375
      Left            =   3120
      Picture         =   "Project1.frx":B62E
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   112
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picIE 
      Height          =   375
      Left            =   3120
      Picture         =   "Project1.frx":BBF0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   111
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picID 
      Height          =   375
      Left            =   3120
      Picture         =   "Project1.frx":C1B2
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   110
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picIC 
      Height          =   375
      Left            =   3120
      Picture         =   "Project1.frx":C774
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   109
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picHC 
      Height          =   375
      Left            =   2760
      Picture         =   "Project1.frx":CD36
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   108
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picGC 
      Height          =   375
      Left            =   2400
      Picture         =   "Project1.frx":D2F8
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   107
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picFC 
      Height          =   375
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   106
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picEC 
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   105
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picDC 
      Height          =   375
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   104
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCC 
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   103
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picBC 
      Height          =   375
      Left            =   600
      Picture         =   "Project1.frx":D8BA
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   102
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAC 
      Height          =   375
      Left            =   240
      Picture         =   "Project1.frx":DE7C
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   101
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picIB 
      Height          =   375
      Left            =   3120
      Picture         =   "Project1.frx":E43E
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   100
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picHB 
      Height          =   375
      Left            =   2760
      Picture         =   "Project1.frx":EA00
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   99
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picGB 
      Height          =   375
      Left            =   2400
      Picture         =   "Project1.frx":EFC2
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   98
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picFB 
      Height          =   375
      Left            =   2040
      Picture         =   "Project1.frx":F584
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   97
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picEB 
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   96
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picDB 
      Height          =   375
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   95
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCB 
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   94
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picBB 
      Height          =   375
      Left            =   600
      Picture         =   "Project1.frx":FB46
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   93
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAB 
      Height          =   375
      Left            =   240
      Picture         =   "Project1.frx":10108
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   92
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picIA 
      Height          =   375
      Left            =   3120
      Picture         =   "Project1.frx":106CA
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   90
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picHA 
      Height          =   375
      Left            =   2760
      Picture         =   "Project1.frx":10C8C
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   89
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picGA 
      Height          =   375
      Left            =   2400
      Picture         =   "Project1.frx":1124E
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   88
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picFA 
      Height          =   375
      Left            =   2040
      Picture         =   "Project1.frx":11810
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   87
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picEA 
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   86
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picDA 
      Height          =   375
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   85
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCA 
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   84
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picBA 
      Height          =   375
      Left            =   600
      Picture         =   "Project1.frx":11DD2
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   83
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdII 
      Height          =   375
      Left            =   3120
      TabIndex        =   82
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdIH 
      Height          =   375
      Left            =   3120
      TabIndex        =   81
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdIG 
      Height          =   375
      Left            =   3120
      TabIndex        =   80
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdIF 
      Height          =   375
      Left            =   3120
      TabIndex        =   79
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdIE 
      Height          =   375
      Left            =   3120
      TabIndex        =   78
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdID 
      Height          =   375
      Left            =   3120
      TabIndex        =   77
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdIC 
      Height          =   375
      Left            =   3120
      TabIndex        =   76
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdIB 
      Height          =   375
      Left            =   3120
      TabIndex        =   75
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdHI 
      Height          =   375
      Left            =   2760
      TabIndex        =   74
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdHH 
      Height          =   375
      Left            =   2760
      TabIndex        =   73
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdHG 
      Height          =   375
      Left            =   2760
      TabIndex        =   72
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdHF 
      Height          =   375
      Left            =   2760
      TabIndex        =   71
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdHE 
      Height          =   375
      Left            =   2760
      TabIndex        =   70
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdHD 
      Height          =   375
      Left            =   2760
      TabIndex        =   69
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdHC 
      Height          =   375
      Left            =   2760
      TabIndex        =   68
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdHB 
      Height          =   375
      Left            =   2760
      TabIndex        =   67
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdGI 
      Height          =   375
      Left            =   2400
      TabIndex        =   66
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdGH 
      Height          =   375
      Left            =   2400
      TabIndex        =   65
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdGG 
      Height          =   375
      Left            =   2400
      TabIndex        =   64
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdGF 
      Height          =   375
      Left            =   2400
      TabIndex        =   63
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdGE 
      Height          =   375
      Left            =   2400
      TabIndex        =   62
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdGD 
      Height          =   375
      Left            =   2400
      TabIndex        =   61
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdGC 
      Height          =   375
      Left            =   2400
      TabIndex        =   60
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdGB 
      Height          =   375
      Left            =   2400
      TabIndex        =   59
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdFI 
      Height          =   375
      Left            =   2040
      TabIndex        =   58
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdFH 
      Height          =   375
      Left            =   2040
      TabIndex        =   57
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdFG 
      Height          =   375
      Left            =   2040
      TabIndex        =   56
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdFF 
      Height          =   375
      Left            =   2040
      TabIndex        =   55
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdFE 
      Height          =   375
      Left            =   2040
      TabIndex        =   54
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdFD 
      Height          =   375
      Left            =   2040
      TabIndex        =   53
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdFC 
      Height          =   375
      Left            =   2040
      TabIndex        =   52
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdFB 
      Height          =   375
      Left            =   2040
      TabIndex        =   51
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdEI 
      Height          =   375
      Left            =   1680
      TabIndex        =   50
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdEH 
      Height          =   375
      Left            =   1680
      TabIndex        =   49
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdEG 
      Height          =   375
      Left            =   1680
      TabIndex        =   48
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdEF 
      Height          =   375
      Left            =   1680
      TabIndex        =   47
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdEE 
      Height          =   375
      Left            =   1680
      TabIndex        =   46
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdED 
      Height          =   375
      Left            =   1680
      TabIndex        =   45
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdEC 
      Height          =   375
      Left            =   1680
      TabIndex        =   44
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdEB 
      Height          =   375
      Left            =   1680
      TabIndex        =   43
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdDI 
      Height          =   375
      Left            =   1320
      TabIndex        =   42
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdDH 
      Height          =   375
      Left            =   1320
      TabIndex        =   41
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdDG 
      Height          =   375
      Left            =   1320
      TabIndex        =   40
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdDF 
      Height          =   375
      Left            =   1320
      TabIndex        =   39
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdDE 
      Height          =   375
      Left            =   1320
      TabIndex        =   38
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox picSmiley 
      Height          =   495
      Left            =   1560
      Picture         =   "Project1.frx":12394
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   37
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdDD 
      Height          =   375
      Left            =   1320
      TabIndex        =   36
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdDC 
      Height          =   375
      Left            =   1320
      TabIndex        =   35
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdDB 
      Height          =   375
      Left            =   1320
      TabIndex        =   34
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdCI 
      Height          =   375
      Left            =   960
      TabIndex        =   33
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdCH 
      Height          =   375
      Left            =   960
      TabIndex        =   32
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdCG 
      Height          =   375
      Left            =   960
      TabIndex        =   31
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdCF 
      Height          =   375
      Left            =   960
      TabIndex        =   30
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdCE 
      Height          =   375
      Left            =   960
      TabIndex        =   29
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdCD 
      Height          =   375
      Left            =   960
      TabIndex        =   28
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdCC 
      Height          =   375
      Left            =   960
      TabIndex        =   27
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdCB 
      Height          =   375
      Left            =   960
      TabIndex        =   26
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdBI 
      Height          =   375
      Left            =   600
      TabIndex        =   25
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdBH 
      Height          =   375
      Left            =   600
      TabIndex        =   24
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdBG 
      Height          =   375
      Left            =   600
      TabIndex        =   23
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdBF 
      Height          =   375
      Left            =   600
      TabIndex        =   22
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdBE 
      Height          =   375
      Left            =   600
      TabIndex        =   21
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdBD 
      Height          =   375
      Left            =   600
      TabIndex        =   20
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdBC 
      Height          =   375
      Left            =   600
      TabIndex        =   19
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdBB 
      Height          =   375
      Left            =   600
      TabIndex        =   18
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdAI 
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdAH 
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdAG 
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdAF 
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdAE 
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdAD 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdAC 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdAB 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdIA 
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdHA 
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdGA 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdFA 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdEA 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdDA 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdCA 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdBA 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox picAA 
      Height          =   375
      Left            =   240
      Picture         =   "Project1.frx":12D06
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   0
         TabIndex        =   91
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdAA 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   375
   End
End
Attribute VB_Name = "frmMinesweeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTR As Integer
'Each of the squares is a button that corresponds to a picture
'depending if the picture is blank or has a number or is a bomb
'it will display a certain number of pictures over the buttons

Private Sub cmdAA_Click()

picAA.Visible = True
picAD.Visible = True
picAH.Visible = True
picEI.Visible = True
picEG.Visible = True
picFG.Visible = True
picGA.Visible = True
picIA.Visible = True
picHD.Visible = True
picIE.Visible = True
cmdAA.Enabled = False
picSmiley.Visible = False
picFrown.Visible = True
MsgBox ("Im sorry, " & Name1 & ", you lost.")

End Sub

Private Sub cmdAB_Click()
If picAB.Visible = False Then
    CTR = CTR + 1
End If
picAB.Visible = True
cmdAB.Enabled = False
End Sub

Private Sub cmdAC_Click()
If picAC.Visible = False Then
    CTR = CTR + 1
End If
picAC.Visible = True
cmdAC.Enabled = False
End Sub

Private Sub cmdAD_Click()

picAA.Visible = True
picAD.Visible = True
picAH.Visible = True
picEI.Visible = True
picEG.Visible = True
picFG.Visible = True
picGA.Visible = True
picIA.Visible = True
picHD.Visible = True
picIE.Visible = True
cmdAD.Enabled = False

picSmiley.Visible = False
picFrown.Visible = True
MsgBox ("Im sorry, " & Name1 & ", you lost.")
End Sub

Private Sub cmdAE_Click()
If picAE.Visible = False Then
    CTR = CTR + 10
End If
picAE.Visible = True
cmdAE.Enabled = False
End Sub

Private Sub cmdAF_Click()
If picAF.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picA7.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True

cmdAF.Enabled = False
End Sub

Private Sub cmdAG_Click()
If picAG.Visible = False Then
    CTR = CTR + 1
End If
picAG.Visible = True
cmdAG.Enabled = False
End Sub

Private Sub cmdAH_Click()
picAA.Visible = True
picAD.Visible = True
picAH.Visible = True
picEI.Visible = True
picEG.Visible = True
picFG.Visible = True
picGA.Visible = True
picIA.Visible = True
picHD.Visible = True
picIE.Visible = True
cmdAH.Enabled = False
MsgBox ("Im sorry, " & Name1 & ", you lost.")
End Sub

Private Sub cmdAI_Click()
If picAI.Visible = False Then
    CTR = CTR + 1
End If
picAI.Visible = True
cmdAI.Enabled = False
End Sub

Private Sub cmdBA_Click()
If picBA.Visible = False Then
    CTR = CTR + 1
End If
picBA.Visible = True
cmdBA.Enabled = False
End Sub

Private Sub cmdBackToMathScreen_Click()

'This takes the user back to the math screen

frmMinesweeper.Hide
frmMath.Show

End Sub

Private Sub cmdBB_Click()
If picBB.Visible = False Then
    CTR = CTR + 1
End If
picBB.Visible = True
cmdBB.Enabled = False
End Sub

Private Sub cmdBC_Click()
If pic.Visible = False Then
    CTR = CTR + 1
End If
picBC.Visible = True
cmdBC.Enabled = False
End Sub

Private Sub cmdBD_Click()
If picBD.Visible = False Then
    CTR = CTR + 1
End If
picBD.Visible = True
cmdBD.Enabled = False
End Sub

Private Sub cmdBE_Click()
If picBE.Visible = False Then
    CTR = CTR + 1
End If
picBE.Visible = True
cmdBE.Enabled = False
End Sub

Private Sub cmdBF_Click()
If picBF.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdBF.Enabled = False
End Sub

Private Sub cmdBG_Click()
If picBG.Visible = False Then
    CTR = CTR + 1
End If
picBG.Visible = True
cmdBG.Enabled = False
End Sub

Private Sub cmdBH_Click()
If picBH.Visible = False Then
    CTR = CTR + 1
End If
picBH.Visible = True
cmdBH.Enabled = False
End Sub

Private Sub cmdBI_Click()
If picBI.Visible = False Then
    CTR = CTR + 1
End If
picBI.Visible = True
cmdBI.Enabled = False
End Sub

Private Sub cmdCA_Click()
If picCA.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdCA.Enabled = False
End Sub

Private Sub cmdCB_Click()
If picCB.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdCB.Enabled = False
End Sub

Private Sub cmdCC_Click()
If picCC.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdCC.Enabled = False
End Sub

Private Sub cmdCD_Click()
If picCD.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdCE.Enabled = False
End Sub

Private Sub cmdCF_Click()
If picCF.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdCF.Enabled = False
End Sub

Private Sub cmdCG_Click()
If picCG.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdCG.Enabled = False
End Sub

Private Sub cmdCH_Click()
If picCH.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdCH.Enabled = False
End Sub

Private Sub cmdCI_Click()
If picCI.Visible = False Then
    CTR = CTR + 1
End If
picCI.Visible = True
cmdCI.Enabled = False
End Sub

Private Sub cmdDA_Click()
If picDA.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdDA.Enabled = False
End Sub

Private Sub cmdDB_Click()
If picDB.Visible = False Then
    CTR = CTR + 1
End If
picDB.Visible = True
cmdDB.Enabled = False
End Sub

Private Sub cmdDC_Click()
If picDC.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdDC.Enabled = False
End Sub

Private Sub cmdDD_Click()
If picDD.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdDD.Enabled = False
End Sub

Private Sub cmdDE_Click()
If picDE.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdDE.Enabled = False
End Sub

Private Sub cmdDF_Click()
If picDF.Visible = False Then
    CTR = CTR + 1
End If
picDF.Visible = True
cmdDF.Enabled = False
End Sub

Private Sub cmdDG_Click()
If picDG.Visible = False Then
    CTR = CTR + 1
End If
picDG.Visible = True
cmdDG.Enabled = False
End Sub

Private Sub cmdDH_Click()
If picDH.Visible = False Then
    CTR = CTR + 1
End If
picDH.Visible = True
cmdDH.Enabled = False
End Sub

Private Sub cmdDI_Click()
If picDI.Visible = False Then
    CTR = CTR + 1
End If
picDI.Visible = True
cmdDI.Enabled = False
End Sub

Private Sub cmdEA_Click()
If picEA.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdEA.Enabled = False
End Sub

Private Sub cmdEB_Click()
If picEB.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdEB.Enabled = False
End Sub

Private Sub cmdEC_Click()
If picEC.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdEC.Enabled = False
End Sub

Private Sub cmdED_Click()
If picED.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdED.Enabled = False
End Sub

Private Sub cmdEE_Click()
If picEE.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdEE.Enabled = False
End Sub

Private Sub cmdEF_Click()
If picEF.Visible = False Then
    CTR = CTR + 1
End If
picEF.Visible = True
cmdEF.Enabled = False
End Sub

Private Sub cmdEG_Click()
picAA.Visible = True
picAD.Visible = True
picAH.Visible = True
picEI.Visible = True
picEG.Visible = True
picFG.Visible = True
picGA.Visible = True
picIA.Visible = True
picHD.Visible = True
picIE.Visible = True
cmdEG.Enabled = False

picSmiley.Visible = False
picFrown.Visible = True
MsgBox ("Im sorry, " & Name1 & ", you lost.")
End Sub

Private Sub cmdEH_Click()
If picEH.Visible = False Then
    CTR = CTR + 1
End If
picEH.Visible = True
cmdEH.Enabled = False
End Sub

Private Sub cmdEI_Click()
picAA.Visible = True
picAD.Visible = True
picAH.Visible = True
picEI.Visible = True
picEG.Visible = True
picFG.Visible = True
picGA.Visible = True
picIA.Visible = True
picHD.Visible = True
picIE.Visible = True
cmdEI.Enabled = False

picSmiley.Visible = False
picFrown.Visible = True
MsgBox ("Im sorry, " & Name1 & ", you lost.")
End Sub

Private Sub cmdFA_Click()
If picFA.Visible = False Then
    CTR = CTR + 1
End If
picFA.Visible = True
cmdFA.Enabled = False
End Sub

Private Sub cmdFB_Click()
If picFB.Visible = False Then
    CTR = CTR + 1
End If
picFB.Visible = True
cmdFB.Enabled = False
End Sub

Private Sub cmdFC_Click()
If picFC.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdFC.Enabled = False
End Sub

Private Sub cmdFD_Click()
If picFD.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdFD.Enabled = False
End Sub

Private Sub cmdFE_Click()
If picFE.Visible = False Then
    CTR = CTR + 47
End If
picCA.Visible = True
picCB.Visible = True
picCC.Visible = True
picCD.Visible = True
picCE.Visible = True
picCF.Visible = True
picCG.Visible = True
picCH.Visible = True
picCI.Visible = True

picBA.Visible = True
picBB.Visible = True
picBC.Visible = True
picBD.Visible = True
picBE.Visible = True
picBF.Visible = True
picBG.Visible = True
picBH.Visible = True
picBI.Visible = True

picAE.Visible = True
picAF.Visible = True
picAG.Visible = True

picDA.Visible = True
picDB.Visible = True
picDC.Visible = True
picDD.Visible = True
picDE.Visible = True
picDF.Visible = True
picDG.Visible = True
picDH.Visible = True
picDI.Visible = True

picEA.Visible = True
picEB.Visible = True
picEC.Visible = True
picED.Visible = True
picEE.Visible = True
picEF.Visible = True

picFA.Visible = True
picFB.Visible = True
picFC.Visible = True
picFD.Visible = True
picFE.Visible = True
picFF.Visible = True

picGB.Visible = True
picGC.Visible = True
picGD.Visible = True
picGE.Visible = True
picGF.Visible = True
cmdFE.Enabled = False
End Sub

Private Sub cmdFF_Click()
If picFF.Visible = False Then
    CTR = CTR + 1
End If
picFF.Visible = True
cmdFF.Enabled = False
End Sub

Private Sub cmdFG_Click()
picAA.Visible = True
picAD.Visible = True
picAH.Visible = True
picEI.Visible = True
picEG.Visible = True
picFG.Visible = True
picGA.Visible = True
picIA.Visible = True
picHD.Visible = True
picIE.Visible = True
cmdFG.Enabled = False

picSmiley.Visible = False
picFrown.Visible = True

MsgBox ("Im sorry, " & Name1 & ", you lost.")
End Sub

Private Sub cmdFH_Click()
If picFH.Visible = False Then
    CTR = CTR + 1
End If
picFH.Visible = True
cmdFH.Enabled = False
End Sub

Private Sub cmdFI_Click()
If picFI.Visible = False Then
    CTR = CTR + 1
End If
picFI.Visible = True
cmdFI.Enabled = False
End Sub

Private Sub cmdGA_Click()
picAA.Visible = True
picAD.Visible = True
picAH.Visible = True
picEI.Visible = True
picEG.Visible = True
picFG.Visible = True
picGA.Visible = True
picIA.Visible = True
picHD.Visible = True
picIE.Visible = True
cmdGA.Enabled = False

picSmiley.Visible = False
picFrown.Visible = True

MsgBox ("Im sorry, " & Name1 & ", you lost.")
End Sub

Private Sub cmdGB_Click()
If picGB.Visible = False Then
    CTR = CTR + 1
End If
picGB.Visible = True
cmdGB.Enabled = False
End Sub

Private Sub cmdGC_Click()
If picGC.Visible = False Then
    CTR = CTR + 1
End If
picGC.Visible = True
cmdGC.Enabled = False
End Sub

Private Sub cmdGD_Click()
If picGD.Visible = False Then
    CTR = CTR + 1
End If
picGD.Visible = True
cmdGD.Enabled = False
End Sub

Private Sub cmdGE_Click()
If picGE.Visible = False Then
    CTR = CTR + 1
End If
picGE.Visible = True
cmdGE.Enabled = False
End Sub

Private Sub cmdGF_Click()
If picGF.Visible = False Then
    CTR = CTR + 1
End If
picGF.Visible = True
cmdGF.Enabled = False
End Sub

Private Sub cmdGG_Click()
If picGG.Visible = False Then
    CTR = CTR + 1
End If
picGG.Visible = True
cmdGG.Enabled = False
End Sub

Private Sub cmdGH_Click()
If picGH.Visible = False Then
    CTR = CTR + 1
End If
picGH.Visible = True
cmdGH.Enabled = False
End Sub

Private Sub cmdGI_Click()
If picGI.Visible = False Then
    CTR = CTR + 10
End If
picAA.Visible = True
picAD.Visible = True
picAH.Visible = True
picEI.Visible = True
picEG.Visible = True
picFG.Visible = True
picGA.Visible = True
picIA.Visible = True
picHD.Visible = True
picIE.Visible = True
cmdGI.Enabled = False
End Sub

Private Sub cmdHA_Click()
If picHA.Visible = False Then
    CTR = CTR + 1
End If
picHA.Visible = True
cmdHA.Enabled = False
End Sub

Private Sub cmdHB_Click()
If picHB.Visible = False Then
    CTR = CTR + 1
End If
picHB.Visible = True
cmdHB.Enabled = False
End Sub

Private Sub cmdHC_Click()
If picHC.Visible = False Then
    CTR = CTR + 1
End If
picHC.Visible = True
cmdHC.Enabled = False
End Sub

Private Sub cmdHD_Click()

picAA.Visible = True
picAD.Visible = True
picAH.Visible = True
picEI.Visible = True
picEG.Visible = True
picFG.Visible = True
picGA.Visible = True
picIA.Visible = True
picHD.Visible = True
picIE.Visible = True
cmdHD.Enabled = False

picSmiley.Visible = False
picFrown.Visible = True
MsgBox ("Im sorry, " & Name1 & ", you lost.")
End Sub

Private Sub cmdHE_Click()
picHE.Visible = True
cmdHE.Enabled = False
End Sub

Private Sub cmdHF_Click()
picHF.Visible = True
cmdHF.Enabled = False
End Sub

Private Sub cmdHG_Click()
picFH.Visible = True
picFI.Visible = True

picGF.Visible = True
picGG.Visible = True
picGH.Visible = True
picGI.Visible = True

picHF.Visible = True
picHG.Visible = True
picHH.Visible = True
picHI.Visible = True

picIF.Visible = True
picIG.Visible = True
picIH.Visible = True
picII.Visible = True

cmdHG.Enabled = False
End Sub

Private Sub cmdHH_Click()
picFH.Visible = True
picFI.Visible = True

picGF.Visible = True
picGG.Visible = True
picGH.Visible = True
picGI.Visible = True

picHF.Visible = True
picHG.Visible = True
picHH.Visible = True
picHI.Visible = True

picIF.Visible = True
picIG.Visible = True
picIH.Visible = True
picII.Visible = True
cmdHH.Enabled = False
End Sub

Private Sub cmdHI_Click()
picFH.Visible = True
picFI.Visible = True

picGF.Visible = True
picGG.Visible = True
picGH.Visible = True
picGI.Visible = True

picHF.Visible = True
picHG.Visible = True
picHH.Visible = True
picHI.Visible = True

picIF.Visible = True
picIG.Visible = True
picIH.Visible = True
picII.Visible = True
cmdHI.Enabled = False
End Sub

Private Sub cmdIA_Click()
picAA.Visible = True
picAD.Visible = True
picAH.Visible = True
picEI.Visible = True
picEG.Visible = True
picFG.Visible = True
picGA.Visible = True
picIA.Visible = True
picHD.Visible = True
picIE.Visible = True
cmdIA.Enabled = False

picSmiley.Visible = False
picFrown.Visible = True

MsgBox ("Im sorry, " & Name1 & ", you lost.")
End Sub

Private Sub cmdIB_Click()
picIB.Visible = True
cmdIB.Enabled = False
End Sub

Private Sub cmdIC_Click()
picIC.Visible = True
cmdIC.Enabled = False
End Sub

Private Sub cmdID_Click()
picID.Visible = True
cmdID.Enabled = False
End Sub

Private Sub cmdIE_Click()
picAA.Visible = True
picAD.Visible = True
picAH.Visible = True
picEI.Visible = True
picEG.Visible = True
picFG.Visible = True
picGA.Visible = True
picIA.Visible = True
picHD.Visible = True
picIE.Visible = True
cmdIE.Enabled = False

picSmiley.Visible = False
picFrown.Visible = True
MsgBox ("Im sorry, " & Name1 & ", you lost.")
End Sub

Private Sub cmdIF_Click()
picIF.Visible = True
cmdIF.Enabled = False
End Sub

Private Sub cmdIG_Click()
picFH.Visible = True
picFI.Visible = True

picGF.Visible = True
picGG.Visible = True
picGH.Visible = True
picGI.Visible = True

picHF.Visible = True
picHG.Visible = True
picHH.Visible = True
picHI.Visible = True

picIF.Visible = True
picIG.Visible = True
picIH.Visible = True
picII.Visible = True
cmdIG.Enabled = False
End Sub

Private Sub cmdIH_Click()
picFH.Visible = True
picFI.Visible = True

picGF.Visible = True
picGG.Visible = True
picGH.Visible = True
picGI.Visible = True

picHF.Visible = True
picHG.Visible = True
picHH.Visible = True
picHI.Visible = True

picIF.Visible = True
picIG.Visible = True
picIH.Visible = True
picII.Visible = True
cmdIH.Enabled = False
End Sub

Private Sub cmdII_Click()
picFH.Visible = True
picFI.Visible = True

picGF.Visible = True
picGG.Visible = True
picGH.Visible = True
picGI.Visible = True

picHF.Visible = True
picHG.Visible = True
picHH.Visible = True
picHI.Visible = True

picIF.Visible = True
picIG.Visible = True
picIH.Visible = True
picII.Visible = True
cmdII.Enabled = False
End Sub

Private Sub cmdQuit_Click()

'This tells the user good luck with their homework and quits

MsgBox ("Good luck with your " & Homework & " hours of homework!")
End
End Sub

Private Sub cmdReset_Click()

'this button resets all the settings so the user can start over

picAA.Visible = False
picAB.Visible = False
picAC.Visible = False
picAD.Visible = False
picAE.Visible = False
picAF.Visible = False
picAG.Visible = False
picAH.Visible = False
picAI.Visible = False

picBA.Visible = False
picBB.Visible = False
picBC.Visible = False
picBD.Visible = False
picBE.Visible = False
picBF.Visible = False
picBG.Visible = False
picBH.Visible = False
picBI.Visible = False

picCA.Visible = False
picCB.Visible = False
picCC.Visible = False
picCD.Visible = False
picCE.Visible = False
picCF.Visible = False
picCG.Visible = False
picCH.Visible = False
picCI.Visible = False

picDA.Visible = False
picDB.Visible = False
picDC.Visible = False
picDD.Visible = False
picDE.Visible = False
picDF.Visible = False
picDG.Visible = False
picDH.Visible = False
picDI.Visible = False

picEA.Visible = False
picEB.Visible = False
picEC.Visible = False
picED.Visible = False
picEE.Visible = False
picEF.Visible = False
picEG.Visible = False
picEH.Visible = False
picEI.Visible = False

picFA.Visible = False
picFB.Visible = False
picFC.Visible = False
picFD.Visible = False
picFE.Visible = False
picFF.Visible = False
picFG.Visible = False
picFH.Visible = False
picFI.Visible = False

picGA.Visible = False
picGB.Visible = False
picGC.Visible = False
picGD.Visible = False
picGE.Visible = False
picGF.Visible = False
picGG.Visible = False
picGH.Visible = False
picGI.Visible = False

picHA.Visible = False
picHB.Visible = False
picHC.Visible = False
picHD.Visible = False
picHE.Visible = False
picHF.Visible = False
picHG.Visible = False
picHH.Visible = False
picHI.Visible = False

picIA.Visible = False
picIB.Visible = False
picIC.Visible = False
picID.Visible = False
picIE.Visible = False
picIF.Visible = False
picIG.Visible = False
picIH.Visible = False
picII.Visible = False

cmdAA.Enabled = True
cmdAB.Enabled = True
cmdAC.Enabled = True
cmdAD.Enabled = True
cmdAE.Enabled = True
cmdAF.Enabled = True
cmdAG.Enabled = True
cmdAH.Enabled = True
cmdAI.Enabled = True

cmdBA.Enabled = True
cmdBB.Enabled = True
cmdBC.Enabled = True
cmdBD.Enabled = True
cmdBE.Enabled = True
cmdBF.Enabled = True
cmdBG.Enabled = True
cmdBH.Enabled = True
cmdBI.Enabled = True

cmdCA.Enabled = True
cmdCB.Enabled = True
cmdCC.Enabled = True
cmdCD.Enabled = True
cmdCE.Enabled = True
cmdCF.Enabled = True
cmdCG.Enabled = True
cmdCH.Enabled = True
cmdCI.Enabled = True

cmdDA.Enabled = True
cmdDB.Enabled = True
cmdDC.Enabled = True
cmdDD.Enabled = True
cmdDE.Enabled = True
cmdDF.Enabled = True
cmdDG.Enabled = True
cmdDH.Enabled = True
cmdDI.Enabled = True

cmdEA.Enabled = True
cmdEB.Enabled = True
cmdEC.Enabled = True
cmdED.Enabled = True
cmdEE.Enabled = True
cmdEF.Enabled = True
cmdEG.Enabled = True
cmdEH.Enabled = True
cmdEI.Enabled = True

cmdFA.Enabled = True
cmdFB.Enabled = True
cmdFC.Enabled = True
cmdFD.Enabled = True
cmdFE.Enabled = True
cmdFF.Enabled = True
cmdFG.Enabled = True
cmdFH.Enabled = True
cmdFI.Enabled = True

cmdGA.Enabled = True
cmdGB.Enabled = True
cmdGC.Enabled = True
cmdGD.Enabled = True
cmdGE.Enabled = True
cmdGF.Enabled = True
cmdGG.Enabled = True
cmdGH.Enabled = True
cmdGI.Enabled = True

cmdHA.Enabled = True
cmdHB.Enabled = True
cmdHC.Enabled = True
cmdHD.Enabled = True
cmdHE.Enabled = True
cmdHF.Enabled = True
cmdHG.Enabled = True
cmdHH.Enabled = True
cmdHI.Enabled = True

cmdIA.Enabled = True
cmdIB.Enabled = True
cmdIC.Enabled = True
cmdID.Enabled = True
cmdIE.Enabled = True
cmdIF.Enabled = True
cmdIG.Enabled = True
cmdIH.Enabled = True
cmdII.Enabled = True
picFrown.Visible = False
picSmiley.Visible = True
End Sub

