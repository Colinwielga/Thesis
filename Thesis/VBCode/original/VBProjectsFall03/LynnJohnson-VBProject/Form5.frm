VERSION 5.00
Begin VB.Form FrmGame4 
   BackColor       =   &H00800080&
   Caption         =   "Game 4"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   10305
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfind 
      Caption         =   "Click Here"
      Height          =   615
      Left            =   10200
      TabIndex        =   41
      Top             =   3960
      Width           =   1335
   End
   Begin VB.PictureBox scoreresults 
      Height          =   495
      Left            =   10560
      ScaleHeight     =   435
      ScaleWidth      =   1755
      TabIndex        =   40
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CommandButton cmdscore 
      Caption         =   "Calculate Score"
      Height          =   855
      Left            =   9360
      TabIndex        =   39
      Top             =   8640
      Width           =   975
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Box"
      Height          =   495
      Left            =   10320
      TabIndex        =   38
      Top             =   6960
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   9840
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   37
      Top             =   4680
      Width           =   2055
   End
   Begin VB.PictureBox pbxresults 
      BackColor       =   &H00FFC0FF&
      Height          =   495
      Left            =   5760
      ScaleHeight     =   435
      ScaleWidth      =   2835
      TabIndex        =   36
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   10800
      TabIndex        =   34
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdplay 
      Caption         =   "Play again"
      Height          =   735
      Left            =   9360
      TabIndex        =   33
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Menu"
      Height          =   735
      Left            =   10080
      TabIndex        =   32
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdthirteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7200
      TabIndex        =   31
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdsixteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4920
      TabIndex        =   30
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdeleven 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2640
      TabIndex        =   29
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdone 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   28
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdtwo 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   27
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdsix 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2640
      TabIndex        =   26
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdnine 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4920
      TabIndex        =   25
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdeight 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7200
      TabIndex        =   24
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdtwelve 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7200
      TabIndex        =   23
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdfive 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4920
      TabIndex        =   22
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdten 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2640
      TabIndex        =   21
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdseven 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   20
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdfour 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   19
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton cmdfifteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2640
      TabIndex        =   18
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton cmdfourteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4920
      TabIndex        =   17
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton cmdthree 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7200
      TabIndex        =   16
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picresults16 
      Height          =   1815
      Left            =   4920
      Picture         =   "Form5.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   15
      Top             =   1680
      Width           =   1935
   End
   Begin VB.PictureBox picresults14 
      Height          =   1815
      Left            =   4920
      Picture         =   "Form5.frx":4C84
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   14
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picresults13 
      Height          =   1815
      Left            =   7200
      Picture         =   "Form5.frx":9912
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   13
      Top             =   1680
      Width           =   1935
   End
   Begin VB.PictureBox picresults5 
      Height          =   1815
      Left            =   4920
      Picture         =   "Form5.frx":E5A0
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   12
      Top             =   6000
      Width           =   1935
   End
   Begin VB.PictureBox picresults11 
      Height          =   1815
      Left            =   2640
      Picture         =   "Form5.frx":137C7
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.PictureBox picresults10 
      Height          =   1815
      Left            =   2640
      Picture         =   "Form5.frx":17F46
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   10
      Top             =   6000
      Width           =   1935
   End
   Begin VB.PictureBox picresults8 
      Height          =   1815
      Left            =   7200
      Picture         =   "Form5.frx":1E94F
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   9
      Top             =   3840
      Width           =   1935
   End
   Begin VB.PictureBox picresults12 
      Height          =   1815
      Left            =   7200
      Picture         =   "Form5.frx":24344
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   8
      Top             =   6000
      Width           =   1935
   End
   Begin VB.PictureBox picresults15 
      Height          =   1815
      Left            =   2640
      Picture         =   "Form5.frx":28AC3
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picresults6 
      Height          =   1815
      Left            =   2640
      Picture         =   "Form5.frx":2D747
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.PictureBox picresults9 
      Height          =   1815
      Left            =   4920
      Picture         =   "Form5.frx":3296E
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   3840
      Width           =   1935
   End
   Begin VB.PictureBox picresults4 
      Height          =   1815
      Left            =   360
      Picture         =   "Form5.frx":39377
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picresults3 
      Height          =   1815
      Left            =   7200
      Picture         =   "Form5.frx":3EB31
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picresults7 
      Height          =   1815
      Left            =   360
      Picture         =   "Form5.frx":442EB
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   6000
      Width           =   1935
   End
   Begin VB.PictureBox picresults2 
      Height          =   1815
      Left            =   360
      Picture         =   "Form5.frx":49CE0
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   3840
      Width           =   1935
   End
   Begin VB.PictureBox picresults1 
      Height          =   1815
      Left            =   360
      Picture         =   "Form5.frx":4EA3F
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblclick 
      Caption         =   "Do you want to know what you're looking for?  Click below and find out!"
      Height          =   495
      Left            =   9480
      TabIndex        =   42
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   $"Form5.frx":5379E
      Height          =   735
      Left            =   480
      TabIndex        =   35
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "FrmGame4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
