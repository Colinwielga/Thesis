VERSION 5.00
Begin VB.Form frmLevel3 
   Caption         =   "Level 3"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8220
   LinkTopic       =   "Form2"
   ScaleHeight     =   6270
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Level!"
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdFavorite 
      Caption         =   "Who is your favorite Character?"
      Height          =   1095
      Left            =   5760
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton cmdCharacter 
      Caption         =   "Get Characters"
      Height          =   1095
      Left            =   5640
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Image imgQuit 
      Height          =   705
      Left            =   7200
      Picture         =   "frmLevel3.frx":0000
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label lblDirections 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   $"frmLevel3.frx":0505
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   6135
   End
   Begin VB.Image imgLevel3 
      Height          =   6675
      Left            =   0
      Picture         =   "frmLevel3.frx":05C5
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   8175
   End
End
Attribute VB_Name = "frmLevel3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Names(1 To 100) As String, Facts(1 To 100) As String, Counter As Integer

Private Sub cmdCharacter_Click()
    Dim CharacterName As String, fact As String
    Open App.Path & "\Characters.txt" For Input As #1       'Opens File into program
    Do Until EOF(1)                                         'Goes through contents of file
        Input #1, CharacterName, fact                       'Inputs contents of file into program
        Counter = Counter + 1                               ' Adds one to counter
        picResults.Print CharacterName                      'Prints contents of file
        Names(Counter) = CharacterName                      'stores contents of file in array
        Facts(Counter) = fact                               'stores contents of file in array
    Loop                                                    ' Loops to Do Until
    Close #1                                                ' Closes text document
    
End Sub

Private Sub cmdFavorite_Click()
    picResults.Cls                  ' Clears picture box
    Dim Found As Boolean, favorite As String, N As Integer
    Found = False                   ' sets found equal to false
    favorite = InputBox("Enter your favorite character from the list!", "Favorite Character")  ' Inputs characters name into file

    Do While (Found = False And N < (Counter))      ' Starts search
        N = N + 1                                   ' Adds 1 to N
        If Names(N) = favorite Then                 ' Compares input to array
            Found = True                            ' Sets Found equal to True
        End If                                      ' Ends If statement
    Loop                                            ' Loops to Do While
    
    If Found = True Then                            ' Results of Do While Loop
        picResults.Print Names(N)                   ' Prints Name
        picResults.Print Facts(N)                   ' Prints Fact
    Else                                            ' If not found
        picResults.Print "Not a Character in the List!" ' Print Error Message
    End If                                          ' End If Statement
        
    
End Sub

Private Sub cmdNext_Click()
    frmLevel3.Visible = False           ' Hide Level 3
    frmLevel4.Visible = True            ' Show Level 4
    
End Sub


Private Sub imgQuit_Click()
    End                                 ' End Program
End Sub
