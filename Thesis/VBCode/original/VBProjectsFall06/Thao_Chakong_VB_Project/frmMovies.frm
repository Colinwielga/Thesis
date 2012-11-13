VERSION 5.00
Begin VB.Form frmMovies 
   BackColor       =   &H80000006&
   Caption         =   "Movie Descriptions"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmMovies.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Back to Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdSortYear 
      Caption         =   "Sort by Year (Ascending)"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortTitle 
      Caption         =   "Sort by Title"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Main Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Movies List"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   8535
      Left            =   5160
      ScaleHeight     =   8475
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   360
      Width           =   7095
   End
End
Attribute VB_Name = "frmMovies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Planet of Jet Li
'Form Name: frmMovies
'Author: Chakong Thao
'Date Written: Monday, Oct. 30th
'Form Objective: This form strictly displays many of Jet Li's
                'movies and also allowing the user to sort it
                'alphabetically or in ascending years.
                
Option Explicit

Private Sub cmdBack_Click() 'This brings user back to General page
    frmMovies.Hide
    frmGeneral.Show
End Sub

Private Sub cmdDisplay_Click()  'This reads the array file of movie titles and displays it in the picture box
    Open App.Path & "\MovieList.txt" For Input As #1
    Counter = 0
    picResults.Cls
    
    Do Until EOF(1)
        Input #1, Movie, Year
        Counter = Counter + 1
        Movies(Counter) = Movie
        Years(Counter) = Year
    Loop
    
    Close #1
    
    picResults.Print "Movie Title"; Tab(50); "Year Made"
    picResults.Print "-----------------------------------------------------------------------------------------------"
    
    For Pos = 1 To Counter
        picResults.Print Movies(Pos); Tab(50); Years(Pos)
    Next Pos
    
End Sub

Private Sub cmdMain_Click() 'This brings user back to beginning page
    frmMovies.Hide
    frmJetLi.Show
End Sub

Private Sub cmdSearch_Click()   'This brings user to page with all available movies for sale
    frmMovies.Hide
    frmMovieSale.Show
End Sub

Private Sub cmdSortTitle_Click()    'Most importantly, this button compares the arrays in the file and stores them alphabetically
    picResults.Cls
    
    For Pass = 1 To Counter - 1
        For Comp = 1 To Counter - Pass
            If Movies(Comp) > Movies(Comp + 1) Then
                TempMovie = Movies(Comp)
                Movies(Comp) = Movies(Comp + 1)
                Movies(Comp + 1) = TempMovie
                
                TempYear = Years(Comp)
                Years(Comp) = Years(Comp + 1)
                Years(Comp + 1) = TempYear
            End If
        Next Comp
    Next Pass
    
    picResults.Print "Movie Title"; Tab(50); "Year Made"
    picResults.Print "-----------------------------------------------------------------------------------------------"
    
    For Pos = 1 To Counter
        picResults.Print Movies(Pos); Tab(50); Years(Pos)
    Next Pos
    
    
        
End Sub

Private Sub cmdSortYear_Click() 'Also very important, this button reads the file and stores the list in ascending year order
    picResults.Cls
    
    For Pass = 1 To Counter - 1
        For Comp = 1 To Counter - Pass
            If Years(Comp) > Years(Comp + 1) Then
                TempYear = Years(Comp)
                Years(Comp) = Years(Comp + 1)
                Years(Comp + 1) = TempYear
                
                TempMovie = Movies(Comp)
                Movies(Comp) = Movies(Comp + 1)
                Movies(Comp + 1) = TempMovie
            End If
        Next Comp
    Next Pass
    
    picResults.Print "Movie Title"; Tab(50); "Year Made"
    picResults.Print "-----------------------------------------------------------------------------------------------"
    
    For Pos = 1 To Counter
        picResults.Print Movies(Pos); Tab(50); Years(Pos)
    Next Pos
End Sub
