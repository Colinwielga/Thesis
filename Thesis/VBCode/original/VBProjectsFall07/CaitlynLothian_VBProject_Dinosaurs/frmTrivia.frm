VERSION 5.00
Begin VB.Form frmTrivia 
   BackColor       =   &H00008000&
   Caption         =   "Learn about Dinosaurs!"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDinos 
      Height          =   3975
      Left            =   8280
      Picture         =   "frmTrivia.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   5475
      TabIndex        =   23
      Top             =   120
      Width           =   5535
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   6120
      TabIndex        =   20
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find My Dino!"
      Height          =   855
      Left            =   6120
      TabIndex        =   19
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   855
      Left            =   6120
      TabIndex        =   18
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack11 
      Caption         =   "Back to main page"
      Height          =   855
      Left            =   6120
      TabIndex        =   17
      Top             =   1080
      Width           =   1935
   End
   Begin VB.PictureBox picResultsInformation 
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   13875
      TabIndex        =   16
      Top             =   4320
      Width           =   13935
   End
   Begin VB.TextBox txtDinosaur 
      Height          =   615
      Left            =   2160
      TabIndex        =   14
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Label lblDinoTable 
      BackColor       =   &H00008000&
      Caption         =   "All Dinosaurs within the database are listed below. Please type in the name EXACTLY as you see it."
      Height          =   615
      Left            =   240
      TabIndex        =   22
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblBibliography 
      BackColor       =   &H00008000&
      Caption         =   $"frmTrivia.frx":85B8
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   480
      TabIndex        =   21
      Top             =   6240
      Width           =   12975
   End
   Begin VB.Label lblName 
      BackColor       =   &H00008000&
      Caption         =   "Please enter the name of the dinosaur you would like to know about"
      Height          =   855
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblDino14 
      BackColor       =   &H00008000&
      Caption         =   "Tyrannosaurus Rex"
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblDino13 
      BackColor       =   &H00008000&
      Caption         =   "Torosaurus"
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblDino12 
      BackColor       =   &H00008000&
      Caption         =   "Suchimimus"
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblDino11 
      BackColor       =   &H00008000&
      Caption         =   "Stegosaurus"
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblDino10 
      BackColor       =   &H00008000&
      Caption         =   "Parasaurolophus"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblDino9 
      BackColor       =   &H00008000&
      Caption         =   "Pachycephasaurus"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblDino8 
      BackColor       =   &H00008000&
      Caption         =   "Omeisaurus"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblDino7 
      BackColor       =   &H00008000&
      Caption         =   "Mullaburrasaurus"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblDino6 
      BackColor       =   &H00008000&
      Caption         =   "Diolophosaurus"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblDino5 
      BackColor       =   &H00008000&
      Caption         =   "Deinonychus (Watcher)"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblDino4 
      BackColor       =   &H00008000&
      Caption         =   "Deinonychus (Slasher)"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblDino3 
      BackColor       =   &H00008000&
      Caption         =   "Chasmosaurus"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblDino2 
      BackColor       =   &H00008000&
      Caption         =   "Baryonyx"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblDino1 
      BackColor       =   &H00008000&
      Caption         =   "Amaragasaurus"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "frmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declare all variables to be used only in this form
    Dim DinoName(1 To 100) As String
    Dim DinoFact(1 To 100) As String
    Dim DinoDiet(1 To 100) As String
    Dim DinoHomeland(1 To 100) As String
    Dim DinoPeriod(1 To 100) As String

Private Sub cmdBack11_Click()
    'When the user clicks this button, the load page appears
    frmTrivia.Visible = False
    frmLoad.Visible = True
End Sub

Private Sub cmdFind_Click()
    'Dim variables to be used only in this subroutine
    Dim SDinoName As String
    Dim Found As Boolean
    Dim Pos As Integer
    Dim CTR As Integer
    
    'Set original values
    Found = False
    Pos = 0
    CTR = 0
    
    'Load the data file
    Open App.Path & "\Dinosaur_Trivia_List.txt" For Input As #1     'Inputs the Array
        Do While Not EOF(1)                                         'Done until the End of the File
            CTR = CTR + 1                                           'Increments by one, so it can keep track of the data
            Input #1, DinoName(CTR), DinoFact(CTR), DinoDiet(CTR), DinoHomeland(CTR), DinoPeriod(CTR)
        Loop
    Close #1
    
   
       
    'Uses Match and Stop to find Dinosaur within list
    SDinoName = txtDinosaur.Text            'User enters desired dinosaur
    Do While Found = False And Pos < CTR
        Pos = Pos + 1
        If SDinoName = DinoName(Pos) Then   'If a match is found, the flag is changed
            Found = True
        End If
    Loop
    
    'Display matching dinosaur information
    If Found = True Then                     'If a match was found, the information is displayed in a picture box, if a match is not found, an error result is printed
        picResultsInformation.Print "Dinosaur Name:"; Tab(20); DinoName(Pos)
        picResultsInformation.Print "Random Fact:"; Tab(20); DinoFact(Pos)
        picResultsInformation.Print "Diet:"; Tab(20); DinoDiet(Pos)
        picResultsInformation.Print "Homeland:"; Tab(20); DinoHomeland(Pos)
        picResultsInformation.Print "Time Period:"; Tab(20); DinoPeriod(Pos)
    Else
        picResultsInformation.Print "Error. There is no matching dinosaur in the database. Please choose another from the list."
    End If
    
End Sub

Private Sub cmdQuit_Click()
    'Ends the program
    End
End Sub

Private Sub cmdReset_Click()
    'When a user clicks this button, the picture boxes and text box will clear
    picResultsInformation.Cls
End Sub
