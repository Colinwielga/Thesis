VERSION 5.00
Begin VB.Form frmGame2 
   BackColor       =   &H00FF8080&
   Caption         =   "Game 2"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12510
   LinkTopic       =   "Form2"
   ScaleHeight     =   10380
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox scoreresults 
      Height          =   495
      Left            =   10440
      ScaleHeight     =   435
      ScaleWidth      =   1875
      TabIndex        =   42
      Top             =   8760
      Width           =   1935
   End
   Begin VB.CommandButton cmdscore 
      Caption         =   "Calculate Score"
      Height          =   735
      Left            =   9360
      TabIndex        =   41
      Top             =   8640
      Width           =   855
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Box"
      Height          =   615
      Left            =   10200
      TabIndex        =   40
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Click Here"
      Height          =   615
      Left            =   10080
      TabIndex        =   39
      Top             =   4320
      Width           =   1575
   End
   Begin VB.PictureBox results 
      Height          =   2055
      Left            =   9840
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   38
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdfifteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   36
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdseven 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2640
      TabIndex        =   35
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdone 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4920
      TabIndex        =   34
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdsix 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7200
      TabIndex        =   33
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdfour 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7200
      TabIndex        =   32
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdthirteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4920
      TabIndex        =   31
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdfourteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2640
      TabIndex        =   30
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdten 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   29
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdeleven 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   28
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdtwo 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2640
      TabIndex        =   27
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdsixteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4920
      TabIndex        =   26
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdeight 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7200
      TabIndex        =   25
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdthree 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7200
      TabIndex        =   24
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdnine 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4920
      TabIndex        =   23
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdtwelve 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2640
      TabIndex        =   22
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdfive 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   21
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   11040
      TabIndex        =   20
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdplay 
      Caption         =   "Play again"
      Height          =   735
      Left            =   9600
      TabIndex        =   19
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Menu"
      Height          =   735
      Left            =   10320
      TabIndex        =   18
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox pbxresults 
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   6240
      ScaleHeight     =   435
      ScaleWidth      =   2955
      TabIndex        =   17
      Top             =   360
      Width           =   3015
   End
   Begin VB.PictureBox picresults10 
      Height          =   1815
      Left            =   360
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   15
      Top             =   5640
      Width           =   1935
   End
   Begin VB.PictureBox picresults8 
      Height          =   1815
      Left            =   7200
      Picture         =   "Form3.frx":6A09
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   14
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox picresults14 
      Height          =   1815
      Left            =   2640
      Picture         =   "Form3.frx":C3FE
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   13
      Top             =   5640
      Width           =   1935
   End
   Begin VB.PictureBox picresults2 
      Height          =   1815
      Left            =   2640
      Picture         =   "Form3.frx":1108C
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   12
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox picresults1 
      Height          =   1815
      Left            =   4920
      Picture         =   "Form3.frx":15DEB
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   11
      Top             =   7800
      Width           =   1935
   End
   Begin VB.PictureBox picresults13 
      Height          =   1815
      Left            =   4920
      Picture         =   "Form3.frx":1AB4A
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   10
      Top             =   5640
      Width           =   1935
   End
   Begin VB.PictureBox picresults7 
      Height          =   1815
      Left            =   2640
      Picture         =   "Form3.frx":1F7D8
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   9
      Top             =   7800
      Width           =   1935
   End
   Begin VB.PictureBox picresults9 
      Height          =   1815
      Left            =   4920
      Picture         =   "Form3.frx":251CD
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin VB.PictureBox picresults16 
      Height          =   1815
      Left            =   4920
      Picture         =   "Form3.frx":2BBD6
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox picresults15 
      Height          =   1815
      Left            =   360
      Picture         =   "Form3.frx":3085A
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   7800
      Width           =   1935
   End
   Begin VB.PictureBox picresults4 
      Height          =   1815
      Left            =   7200
      Picture         =   "Form3.frx":354DE
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   5640
      Width           =   1935
   End
   Begin VB.PictureBox picresults3 
      Height          =   1815
      Left            =   7200
      Picture         =   "Form3.frx":3AC98
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.PictureBox picresults12 
      Height          =   1815
      Left            =   2640
      Picture         =   "Form3.frx":40452
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.PictureBox picresults11 
      Height          =   1815
      Left            =   360
      Picture         =   "Form3.frx":44BD1
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox picresults6 
      Height          =   1815
      Left            =   7200
      Picture         =   "Form3.frx":49350
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   7800
      Width           =   1935
   End
   Begin VB.PictureBox picresults5 
      Height          =   1815
      Left            =   360
      Picture         =   "Form3.frx":4E577
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblclick 
      Caption         =   "Do you want to know what you're looking for?  Click below and find out!"
      Height          =   495
      Left            =   9480
      TabIndex        =   37
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"Form3.frx":5379E
      Height          =   615
      Left            =   840
      TabIndex        =   16
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmGame2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
