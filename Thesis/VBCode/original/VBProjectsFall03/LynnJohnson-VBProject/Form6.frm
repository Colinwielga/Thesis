VERSION 5.00
Begin VB.Form FrmGame5 
   BackColor       =   &H00000080&
   Caption         =   "Game 5"
   ClientHeight    =   10320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfind 
      Caption         =   "Click Here"
      Height          =   615
      Left            =   9720
      TabIndex        =   41
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Box"
      Height          =   495
      Left            =   9840
      TabIndex        =   40
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdscore 
      Caption         =   "Calculate Score"
      Height          =   855
      Left            =   9000
      TabIndex        =   39
      Top             =   8760
      Width           =   975
   End
   Begin VB.PictureBox scoreresults 
      Height          =   495
      Left            =   10200
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   38
      Top             =   8880
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   9480
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   37
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   10680
      TabIndex        =   36
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdplay 
      Caption         =   "Play again"
      Height          =   735
      Left            =   9240
      TabIndex        =   35
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Menu"
      Height          =   735
      Left            =   9960
      TabIndex        =   34
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox pbxresults 
      BackColor       =   &H00C0C0FF&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   3315
      TabIndex        =   33
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton cmdtwo 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   6840
      TabIndex        =   31
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton cmdfifteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4680
      TabIndex        =   30
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton cmdten 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2520
      TabIndex        =   29
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton cmdeleven 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   28
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton cmdnine 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   27
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdthree 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2520
      TabIndex        =   26
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdthirteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4680
      TabIndex        =   25
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdfive 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   6840
      TabIndex        =   24
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdone 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   6840
      TabIndex        =   23
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdseven 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4680
      TabIndex        =   22
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdsix 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2520
      TabIndex        =   21
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdsixteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   20
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdeight 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   19
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmdfourteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2520
      TabIndex        =   18
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmdtwelve 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4680
      TabIndex        =   17
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmdfour 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   6840
      TabIndex        =   16
      Top             =   2040
      Width           =   1935
   End
   Begin VB.PictureBox picresults16 
      Height          =   1815
      Left            =   360
      Picture         =   "Form6.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   15
      Top             =   4080
      Width           =   1935
   End
   Begin VB.PictureBox picresults12 
      Height          =   1815
      Left            =   4680
      Picture         =   "Form6.frx":4C84
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   14
      Top             =   2040
      Width           =   1935
   End
   Begin VB.PictureBox picresults8 
      Height          =   1815
      Left            =   360
      Picture         =   "Form6.frx":9403
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   13
      Top             =   2040
      Width           =   1935
   End
   Begin VB.PictureBox picresults15 
      Height          =   1815
      Left            =   4680
      Picture         =   "Form6.frx":EDF8
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   12
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picresults7 
      Height          =   1815
      Left            =   4680
      Picture         =   "Form6.frx":13A7C
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   11
      Top             =   4080
      Width           =   1935
   End
   Begin VB.PictureBox picresults14 
      Height          =   1815
      Left            =   2520
      Picture         =   "Form6.frx":19471
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   10
      Top             =   2040
      Width           =   1935
   End
   Begin VB.PictureBox picresults11 
      Height          =   1815
      Left            =   360
      Picture         =   "Form6.frx":1E0FF
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   9
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picresults10 
      Height          =   1815
      Left            =   2520
      Picture         =   "Form6.frx":2287E
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   8
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picresults4 
      Height          =   1815
      Left            =   6840
      Picture         =   "Form6.frx":29287
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   2040
      Width           =   1935
   End
   Begin VB.PictureBox picresults3 
      Height          =   1815
      Left            =   2520
      Picture         =   "Form6.frx":2EA41
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   6120
      Width           =   1935
   End
   Begin VB.PictureBox picresults13 
      Height          =   1815
      Left            =   4680
      Picture         =   "Form6.frx":341FB
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   6120
      Width           =   1935
   End
   Begin VB.PictureBox picresults6 
      Height          =   1815
      Left            =   2520
      Picture         =   "Form6.frx":38E89
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   4080
      Width           =   1935
   End
   Begin VB.PictureBox picresults9 
      Height          =   1815
      Left            =   360
      Picture         =   "Form6.frx":3E0B0
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   6120
      Width           =   1935
   End
   Begin VB.PictureBox picresults5 
      Height          =   1815
      Left            =   6840
      Picture         =   "Form6.frx":44AB9
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   6120
      Width           =   1935
   End
   Begin VB.PictureBox picresults2 
      Height          =   1815
      Left            =   6840
      Picture         =   "Form6.frx":49CE0
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picresults1 
      Height          =   1815
      Left            =   6840
      Picture         =   "Form6.frx":4EA3F
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label lblclick 
      Caption         =   "Do you want to know what you're looking for?  Click below and find out!"
      Height          =   495
      Left            =   9240
      TabIndex        =   42
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   $"Form6.frx":5379E
      Height          =   735
      Left            =   360
      TabIndex        =   32
      Top             =   480
      Width           =   4815
   End
End
Attribute VB_Name = "FrmGame5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
