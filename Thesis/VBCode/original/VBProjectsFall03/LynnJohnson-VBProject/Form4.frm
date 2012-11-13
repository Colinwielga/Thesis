VERSION 5.00
Begin VB.Form frmGame3 
   BackColor       =   &H00008000&
   Caption         =   "Game 3"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   ScaleHeight     =   10305
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Box"
      Height          =   495
      Left            =   10440
      TabIndex        =   42
      Top             =   6720
      Width           =   1335
   End
   Begin VB.PictureBox scoreresults 
      Height          =   495
      Left            =   10680
      ScaleHeight     =   435
      ScaleWidth      =   1755
      TabIndex        =   41
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton cmdscore 
      Caption         =   "Calculate Score"
      Height          =   855
      Left            =   9600
      TabIndex        =   40
      Top             =   8520
      Width           =   855
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Click Here"
      Height          =   615
      Left            =   10440
      TabIndex        =   38
      Top             =   3720
      Width           =   1335
   End
   Begin VB.PictureBox results 
      Height          =   2055
      Left            =   10080
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   37
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   11040
      TabIndex        =   36
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdplay 
      Caption         =   "Play again"
      Height          =   735
      Left            =   9600
      TabIndex        =   35
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Menu"
      Height          =   735
      Left            =   10320
      TabIndex        =   34
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdsixteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7440
      TabIndex        =   33
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton cmdthree 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   5160
      TabIndex        =   32
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton cmdone 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2880
      TabIndex        =   31
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton cmdthirteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   600
      TabIndex        =   30
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton cmdfive 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7440
      TabIndex        =   29
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdeleven 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   5160
      TabIndex        =   28
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdten 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2880
      TabIndex        =   27
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdseven 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   600
      TabIndex        =   26
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdeight 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7440
      TabIndex        =   25
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmdnine 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   5160
      TabIndex        =   24
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmdfifteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2880
      TabIndex        =   23
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmdtwelve 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   600
      TabIndex        =   22
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmdfour 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7440
      TabIndex        =   21
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdsix 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   5160
      TabIndex        =   20
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdtwo 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2880
      TabIndex        =   19
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdfourteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   600
      TabIndex        =   18
      Top             =   1440
      Width           =   1935
   End
   Begin VB.PictureBox picresults14 
      Height          =   1815
      Left            =   600
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   17
      Top             =   1440
      Width           =   1935
   End
   Begin VB.PictureBox picresults13 
      Height          =   1815
      Left            =   600
      Picture         =   "Form4.frx":4C8E
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   16
      Top             =   7920
      Width           =   1935
   End
   Begin VB.PictureBox picresults5 
      Height          =   1815
      Left            =   7440
      Picture         =   "Form4.frx":991C
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   15
      Top             =   5760
      Width           =   1935
   End
   Begin VB.PictureBox picresults16 
      Height          =   1815
      Left            =   7440
      Picture         =   "Form4.frx":EB43
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   14
      Top             =   7920
      Width           =   1935
   End
   Begin VB.PictureBox picresults7 
      Height          =   1815
      Left            =   600
      Picture         =   "Form4.frx":137C7
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   13
      Top             =   5760
      Width           =   1935
   End
   Begin VB.PictureBox picresults10 
      Height          =   1815
      Left            =   2880
      Picture         =   "Form4.frx":191BC
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   12
      Top             =   5760
      Width           =   1935
   End
   Begin VB.PictureBox picresults12 
      Height          =   1815
      Left            =   600
      Picture         =   "Form4.frx":1FBC5
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   11
      Top             =   3600
      Width           =   1935
   End
   Begin VB.PictureBox picresults3 
      Height          =   1815
      Left            =   5160
      Picture         =   "Form4.frx":24344
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   10
      Top             =   7920
      Width           =   1935
   End
   Begin VB.PictureBox picresults15 
      Height          =   1815
      Left            =   2880
      Picture         =   "Form4.frx":29AFE
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   9
      Top             =   3600
      Width           =   1935
   End
   Begin VB.PictureBox picresults6 
      Height          =   1815
      Left            =   5160
      Picture         =   "Form4.frx":2E782
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
   End
   Begin VB.PictureBox picresults8 
      Height          =   1815
      Left            =   7440
      Picture         =   "Form4.frx":339A9
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   3600
      Width           =   1935
   End
   Begin VB.PictureBox picresults4 
      Height          =   1815
      Left            =   7440
      Picture         =   "Form4.frx":3939E
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
   Begin VB.PictureBox picresults9 
      Height          =   1815
      Left            =   5160
      Picture         =   "Form4.frx":3EB58
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   3600
      Width           =   1935
   End
   Begin VB.PictureBox picresults11 
      Height          =   1815
      Left            =   5160
      Picture         =   "Form4.frx":45561
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   5760
      Width           =   1935
   End
   Begin VB.PictureBox picresults2 
      Height          =   1815
      Left            =   2880
      Picture         =   "Form4.frx":49CE0
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.PictureBox picresults1 
      Height          =   1815
      Left            =   2880
      Picture         =   "Form4.frx":4EA3F
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   7920
      Width           =   1935
   End
   Begin VB.PictureBox pbxresults 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   435
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblclick 
      Caption         =   "Do you want to know what you're looking for?  Click below and find out!"
      Height          =   495
      Left            =   9720
      TabIndex        =   39
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   $"Form4.frx":5379E
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "frmGame3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
