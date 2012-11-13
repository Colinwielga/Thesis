VERSION 5.00
Begin VB.Form frmHistory 
   BackColor       =   &H003D30AD&
   Caption         =   "History"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13035
   FillColor       =   &H003D30AD&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   13035
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   600
      Picture         =   "frmHistory.frx":0000
      ScaleHeight     =   3075
      ScaleWidth      =   4035
      TabIndex        =   9
      Top             =   1200
      Width           =   4095
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back to Home Page"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10800
      TabIndex        =   7
      Top             =   6720
      Width           =   2055
   End
   Begin VB.TextBox txtFind 
      Height          =   855
      Left            =   4440
      TabIndex        =   5
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Who Chose this Character?"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7080
      TabIndex        =   4
      Top             =   6600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdReverse 
      Caption         =   "Reverse Alphabetically"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdResults 
      Caption         =   "See History"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      TabIndex        =   2
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "Sort Alphabetically"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      Height          =   4455
      Left            =   5520
      ScaleHeight     =   4395
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H003D30AD&
      Caption         =   "Quiz History"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084C11E&
      Height          =   855
      Left            =   600
      TabIndex        =   8
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label lblFind 
      BackColor       =   &H003D30AD&
      Caption         =   "Enter a Character Name:      Belle                  Beast                Gaston              Lumiere"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      TabIndex        =   6
      Top             =   6600
      Width           =   4095
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declare 2 arrays, 1 for Person's name,1 for Character name as strings and a counter
    Dim Person(1 To 20) As String
    Dim Character(1 To 20) As String
    Dim Ctr As Integer

Private Sub cmdAlpha_Click()
'declare variables for sorting
picResults.Cls
    Dim Pass As Integer
    Dim pos As Integer
    Dim Temp As String
    Dim I As Integer
    
            
    For Pass = 1 To Ctr - 1
        For pos = 1 To Ctr - Pass
            If Person(pos) > Person(pos + 1) Then
                Temp = Person(pos)
                Person(pos) = Person(pos + 1)
                Person(pos + 1) = Temp
                Temp = Character(pos)
                Character(pos) = Character(pos + 1)
                Character(pos + 1) = Temp
            End If
        Next pos
    Next Pass
   picResults.Print "Alphabetical Order"
    picResults.Print "***********************"
    For I = 1 To Ctr
        picResults.Print Person(I), Character(I)
    Next I
                    
End Sub



Private Sub cmdGoBack_Click()
frmHistory.Hide
frmPersonality.Show

End Sub

Private Sub cmdResults_Click()
    Dim I As Integer
    cmdSearch.Visible = True
    cmdAlpha.Visible = True
    cmdReverse.Visible = True
    'load file into parallel arrays
    Open App.Path & "\project_bandb.txt" For Input As #1
    
    Ctr = 0
    
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, Person(Ctr), Character(Ctr)
    Loop
    ' print title
     picResults.Print "Quiz History"
     picResults.Print "*****************"
    'print list of people and characters
    For I = 1 To Ctr
        picResults.Print I; "."; Person(I), Character(I)
    Next I
    cmdResults.Visible = False
    

End Sub

Private Sub cmdReverse_Click()
 picResults.Cls
    Dim Pass As Integer
    Dim pos As Integer
    Dim Temp As String
    Dim I As Integer
    
    'sort arrays in reverse alphabetical order
    For Pass = 1 To Ctr - 1
        For pos = 1 To Ctr - Pass
            If Person(pos) < Person(pos + 1) Then
                Temp = Person(pos)
                Person(pos) = Person(pos + 1)
                Person(pos + 1) = Temp
                Temp = Character(pos)
                Character(pos) = Character(pos + 1)
                Character(pos + 1) = Temp
            End If
        Next pos
    Next Pass
    'print results
    picResults.Print "Reverse Alphabetical Order"
    picResults.Print "************************************"
        
    For I = 1 To Ctr
        picResults.Print I; "."; Person(I), Character(I)
    Next I
End Sub

               
Private Sub cmdSearch_Click()
    Dim pos As Integer
    Dim SearchCharacter As String
    Dim Found As Boolean
    'display persons from the arrays that chose the character that the user enters in a textbox
    SearchCharacter = txtFind.Text
    picResults.Cls
    picResults.Print "People who choose this character"
    picResults.Print "***************************************"
    Found = False
    
    For pos = 1 To Ctr
        
    If UCase(Character(pos)) = UCase(SearchCharacter) Then
            Found = True
             picResults.Print Person(pos), Character(pos)
        End If
    Next pos
    'display "error" message if the character doesn't exist
        If Found = False Then
        MsgBox "This character does not exist!"
        End If
        
        
    
End Sub


