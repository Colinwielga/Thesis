VERSION 5.00
Begin VB.Form frmTheFinerThingsInLife 
   BackColor       =   &H00400000&
   Caption         =   "The Finer Things In Life"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSkip 
      Caption         =   "Skip ahead to Summary Page"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   10
      Top             =   9720
      Width           =   5535
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13200
      TabIndex        =   9
      Top             =   9960
      Width           =   1575
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   8
      Top             =   9960
      Width           =   1575
   End
   Begin VB.CommandButton cmdLucky 
      Caption         =   "Are You Feeling Lucky"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   11160
      TabIndex        =   5
      Top             =   5640
      Width           =   3015
   End
   Begin VB.CommandButton cmdCars 
      Caption         =   "Find The Perfect Ride"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5520
      TabIndex        =   4
      Top             =   5880
      Width           =   3255
   End
   Begin VB.CommandButton cmdHouse 
      Caption         =   "Find Your Dream House"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   600
      TabIndex        =   3
      Top             =   5760
      Width           =   3135
   End
   Begin VB.PictureBox picLottery 
      Height          =   3375
      Left            =   9720
      Picture         =   "frmTheFinerThingsInLife.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   4755
      TabIndex        =   2
      Top             =   1920
      Width           =   4815
   End
   Begin VB.PictureBox picCars 
      Height          =   3375
      Left            =   4560
      Picture         =   "frmTheFinerThingsInLife.frx":4C086
      ScaleHeight     =   3315
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   1920
      Width           =   4815
   End
   Begin VB.PictureBox picMansion 
      Height          =   3375
      Left            =   120
      Picture         =   "frmTheFinerThingsInLife.frx":5117E
      ScaleHeight     =   3315
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label lblB 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Select a button to see what the future holds for your dream house, your perfect car, and your chances at winning the lottery!"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1815
      Left            =   1920
      TabIndex        =   7
      Top             =   7680
      Width           =   12015
   End
   Begin VB.Label lblA 
      BackColor       =   &H00400000&
      Caption         =   "Living It Up"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   5280
      TabIndex        =   6
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "frmTheFinerThingsInLife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: The Game Of Life
'Form Name: frmTheFinerThingsInLife
'Authors: Pam Arndt and Alisa Davis
'Date Written: 3/4/09
'Objective: Provides directions to forms to choose a house, a car, or to play a lottery game.
Option Explicit

Private Sub cmdCars_Click() 'to pick cars
frmCars.Show
frmTheFinerThingsInLife.Hide
End Sub

Private Sub cmdHome_Click() 'to return to beginning page
frmBeginning.Show
frmTheFinerThingsInLife.Hide

End Sub

Private Sub cmdHouse_Click() 'to pick a house
frmHouses.Show
frmTheFinerThingsInLife.Hide

End Sub

Private Sub cmdLucky_Click() 'to play the lotto game
frmLucky.Show
frmTheFinerThingsInLife.Hide
End Sub

Private Sub cmdQuit_Click() 'ends program
End
End Sub

Private Sub cmdSkip_Click() 'continues to summary page without picking house, car, or playing lotto game
frmTheFinerThingsInLife.Hide
frmSummary.Show

End Sub
