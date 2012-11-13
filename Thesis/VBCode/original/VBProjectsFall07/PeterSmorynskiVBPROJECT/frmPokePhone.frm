VERSION 5.00
Begin VB.Form frmDestination 
   BackColor       =   &H8000000D&
   Caption         =   "PokePhone Application"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11130
   Begin VB.CommandButton cmdNext 
      Caption         =   "More PokePhone Options"
      Height          =   1095
      Left            =   5880
      TabIndex        =   4
      Top             =   6120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdNear 
      Caption         =   "Miles Search"
      Height          =   2175
      Left            =   1200
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "QUIT!"
      Height          =   1095
      Left            =   8760
      TabIndex        =   2
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Map Message Data"
      Height          =   2295
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      Height          =   5535
      Left            =   5880
      ScaleHeight     =   5475
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   8760
      Left            =   0
      Picture         =   "frmPokePhone.frx":0000
      Top             =   0
      Width           =   11160
   End
End
Attribute VB_Name = "frmDestination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Inspired by Lab 11 and HW3 Ques2
Private Sub Picture1_Click()
picResultsUser.Print Username
End Sub

Private Sub cmdLoad_Click() 'loads data from file array
    
Open App.Path & "\Destination.txt" For Input As #1

CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
    Input #1, Place(CTR), Miles(CTR)
    Loop
       
 Close #1
       
       cmdLoad.Enabled = False
       cmdNear.Enabled = True
End Sub
Private Sub cmdNear_Click() ' search for first data out of range from user input
picResults.Cls
Dim OutRange As Single
    MilesSearch = InputBox("Search for First City Beyond Your Travel Range", "Enter Maximum Travel Range", MilesSearch) 'Search for milage via an input box)
    Pos = 0

    Found = False
    
    Do While Found = False And Pos < CTR                'Boolean Do While statement and loop
        Pos = Pos + 1
    If MilesSearch <= Miles(Pos) Then
        Found = True
    End If
    
    Loop
    
    If Found = False Then
        MsgBox ("Out of Map Data Range! *Please Try Again*")
    End If
    If Found = True And MilesSearch < Miles(Pos) Then
        OutRange = Miles(Pos) - MilesSearch
        picResults.Print Place(Pos); ": "; Int(Miles(Pos)); "miles from current position"
        picResults.Print Place(Pos); " is"; Int(OutRange); "miles outside your desired travel range!"
    End If
    If Found = True And MilesSearch = Miles(Pos) Then
        picResults.Print "* "; Place(Pos); " is the last city within your desired travel range!"
                                                
End If
    cmdNext.Visible = True
End Sub
Private Sub cmdNumerical_Click()
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
            picResults.Print Place(comp); Tab(20); Int(Miles(comp)); Tab(30)
        Next comp
End Sub

Private Sub cmdNext_Click()
frmDestination.Hide
frmDestination2.Show
MsgBox ("Let's use a default Map application to list all cities in PokeSatellite range in alphabetical order, and then we sort them by distance to see what might be WITHIN your desired range of travel."), , ("INSTRUCTION: POKEPHONE APPLICATION/TASK")
End Sub
Private Sub cmdQuit_Click()
End
End Sub

