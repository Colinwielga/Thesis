VERSION 5.00
Begin VB.Form frmDestination2 
   BackColor       =   &H8000000D&
   Caption         =   "PokePhone Application"
   ClientHeight    =   8040
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "QUIT!"
      Height          =   1095
      Left            =   9600
      TabIndex        =   4
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton cmdRtnHub 
      Caption         =   "Return to Pokemon Central!"
      Height          =   1095
      Left            =   6360
      TabIndex        =   3
      Top             =   6720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdNumerical 
      Caption         =   "See Listings by Numerical Distance (in your range)"
      Height          =   1335
      Left            =   3360
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "See All City Listings (Alphabetical) "
      Height          =   1455
      Left            =   3360
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      Height          =   5535
      Left            =   6360
      ScaleHeight     =   5475
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   0
      Picture         =   "frmDestination2.frx":0000
      Top             =   0
      Width           =   9600
   End
   Begin VB.Label lblheaderDest2 
      Caption         =   "CITY NAME              MILES AWAY"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   6255
   End
End
Attribute VB_Name = "frmDestination2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  'inspired by lab 11 and HW3 ques2
Private Sub cmdSort_Click() ' sorts array data alphabetically
picResults.Cls
Dim pass As Integer
Dim comp As Integer
Dim temp As String

    For pass = 1 To CTR - 1

    For comp = 1 To CTR - pass
        If Place(comp) > Place(comp + 1) Then
            temp = Place(comp)
            Place(comp) = Place(comp + 1)
            Place(comp + 1) = temp
            temp = Miles(comp)
            Miles(comp) = Miles(comp + 1)
            Miles(comp + 1) = temp
        End If
        Next comp
        Next pass
        For comp = 1 To CTR
            picResults.Print Place(comp); Tab(20); Int(Miles(comp)); Tab(30)
        Next comp
    cmdSort.Visible = False
    cmdNumerical.Visible = True
    lblheaderDest2.Visible = True
End Sub
Private Sub cmdNumerical_Click() 'sorts by number using array data and the user's input range
picResults.Cls
Dim pass As Integer
Dim comp As Integer
Dim temp As String

    For pass = 1 To CTR - 1

    For comp = 1 To CTR - pass
        If Miles(comp) > Miles(comp + 1) Then
            temp = Place(comp)
            Place(comp) = Place(comp + 1)
            Place(comp + 1) = temp
            temp = Miles(comp)
            Miles(comp) = Miles(comp + 1)
            Miles(comp + 1) = temp
        End If
        Next comp
        Next pass
        For comp = 1 To CTR
        If MilesSearch >= Miles(comp) Then
            picResults.Print Place(comp); Tab(20); Int(Miles(comp)); Tab(30)
        End If
        Next comp

cmdNumerical.Visible = False
cmdRtnHub.Visible = True
End Sub
Private Sub cmdRtnHub_Click() ' return to pokemon central
frmDestination2.Hide
frmCentralHub.Show
MsgBox ("Welcome back to Pokemon Central! That PokePhone was a nifty gadget, eh? Trainers find all sorts of uses for it."), , ("INSTRUCTION: CHOOSE YOUR NEXT DESTINATION!")
End Sub
Private Sub cmdQuit_Click() 'quit program
End
End Sub

