VERSION 5.00
Begin VB.Form frmMovies 
   BackColor       =   &H00FF0000&
   Caption         =   "Disney Movies"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   Picture         =   "frmMovies.frx":0000
   ScaleHeight     =   8025
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search for your Favorite Movie"
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdYear 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sort List by Year Released"
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdAlpha 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sort List Alphabetically"
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H00FFFFC0&
      Caption         =   "See List of Disney Movies"
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   7215
      Left            =   5160
      ScaleHeight     =   7155
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Return to Disney Castle"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   1095
   End
End
Attribute VB_Name = "frmMovies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Disney Project
'Movies
'Lori Nohner
'Written- March 17, 2008
'Objective-  This form allows the user to look at a list of Disney movies.
    'They can sort the list alphabetically or by what year the movie was released.
    'They can also search for a movie with in the list.
    
Option Explicit 'declares variables to the correct data tyes
Dim CTR As Integer
Dim Movies(1 To 33) As String
Dim Year(1 To 33) As Integer
Dim N As Integer

Private Sub cmdAlpha_Click()
    Dim Pass As Integer, Pos As Integer, Temp As String, Temp2 As Integer
    
    
    picResults.Cls 'clears picture box
    For Pass = 1 To CTR - 1 'sorts the movies alphabetically
        For Pos = 1 To CTR - Pass
            If Movies(Pos) > Movies(Pos + 1) Then
                Temp = Movies(Pos)
                Movies(Pos) = Movies(Pos + 1)
                Movies(Pos + 1) = Temp
                    Temp2 = Year(Pos)
                    Year(Pos) = Year(Pos + 1)
                    Year(Pos + 1) = Temp2
            End If
        Next Pos
    Next Pass
        
        'Theses next steps print out a header and the list of movies in alphabetical order
        
    picResults.Print "Name of Movie"; Tab(40); "Year Premiered"
    picResults.Print "******************************************************************"
    For N = 1 To CTR
        picResults.Print Movies(N); Tab(45); Year(N)
    Next N
End Sub

Private Sub cmdDisplay_Click()
    picResults.Cls
    CTR = 0
    
    'reads a ist of movies and the years they were released form a text file
    Open App.Path & "\Movies.txt" For Input As #1 'opens and reads Movies.txt file
        Do Until EOF(1)
            CTR = CTR + 1
            Input #1, Movies(CTR), Year(CTR)
            
        Loop
    Close #1
    
    'prints out list of movies
    picResults.Print "Name of Movie"; Tab(40); "Year Released"
    picResults.Print "******************************************************************"
    For N = 1 To CTR
        picResults.Print Movies(N); Tab(45); Year(N)
    Next N
End Sub

Private Sub cmdExit_Click()
    End 'quits program
End Sub

Private Sub cmdReturn_Click()
    frmMovies.Hide 'hides movies page
    frmDisneyCastle.Show 'returns to disney home page
End Sub

Private Sub cmdSearch_Click()
    Dim Found As Boolean
    Dim FavMovie As String
    
    picResults.Cls
    CTR = 0
    Found = False
    
    FavMovie = InputBox("Enter your favorite Disney movie", "Enter Movie") 'shows user an input box that asks for the users favorite Disney movie
    
    'searches for the movie entered
    Do While (Found = False And CTR < 33)
        CTR = CTR + 1
        If LCase(Movies(CTR)) = LCase(FavMovie) Then
            Found = True
        End If
    Loop
    
    'If the movie is on the list, print out name of movie and year it was released in the picture box.
    If Found = True Then
        picResults.Print "The movie "; Movies(CTR); " was first released in "; Year(CTR)
    Else
        MsgBox "Your favorite movie, " & FavMovie & ", is not on the list. Enter another movie.", , "Error" 'if not found it shows the user a message box saying the movies hasn't been found
    End If
    
End Sub

Private Sub cmdYear_Click()
    
    Dim Pass As Integer, Pos As Integer, Temp As Integer, Temp2 As String
    
    'sorts the list of movies by year they were released
    picResults.Cls
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Year(Pos) > Year(Pos + 1) Then
                Temp = Year(Pos)
                Year(Pos) = Year(Pos + 1)
                Year(Pos + 1) = Temp
                    Temp2 = Movies(Pos)
                    Movies(Pos) = Movies(Pos + 1)
                    Movies(Pos + 1) = Temp2
            End If
        Next Pos
    Next Pass
    
    'prints list of movie sorted by year
    picResults.Print "Name of Movie"; Tab(40); "Year Premiered"
    picResults.Print "******************************************************************"
    For N = 1 To CTR
        picResults.Print Movies(N); Tab(45); Year(N)
    Next N
End Sub

Private Sub picResults_Click()

End Sub
