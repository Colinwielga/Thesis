VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H000000FF&
   Caption         =   "Video Game Search"
   ClientHeight    =   9615
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14895
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   14895
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   1212
      Left            =   2280
      Picture         =   "frmGallery.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   11
      Top             =   7920
      Width           =   1332
   End
   Begin VB.CommandButton cmdGamesSold 
      Caption         =   "Sort According to Number of Games Sold (In Millions)"
      Height          =   732
      Left            =   360
      TabIndex        =   7
      Top             =   5760
      Width           =   2532
   End
   Begin VB.CommandButton cmdGame 
      Caption         =   "Sort According to Game"
      Height          =   732
      Left            =   360
      TabIndex        =   6
      Top             =   4080
      Width           =   2532
   End
   Begin VB.CommandButton cmdSystem 
      Caption         =   "Sort According to System "
      Height          =   732
      Left            =   360
      TabIndex        =   5
      Top             =   4920
      Width           =   2532
   End
   Begin VB.CommandButton cmdRelease 
      Caption         =   "Sort According to Release Date (Month)"
      Height          =   732
      Left            =   360
      TabIndex        =   4
      Top             =   6600
      Width           =   2532
   End
   Begin VB.CommandButton cmdRank 
      Caption         =   "Sort According to Rank "
      Height          =   732
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   2532
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Display the Data"
      Height          =   732
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   2532
   End
   Begin VB.PictureBox picresults 
      Height          =   4932
      Left            =   3120
      ScaleHeight     =   4875
      ScaleWidth      =   11355
      TabIndex        =   1
      Top             =   2400
      Width           =   11412
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return back to the main page"
      Height          =   732
      Left            =   11400
      TabIndex        =   0
      Top             =   8400
      Width           =   3012
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "By Bill Macy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   9240
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "The top 20 Video Games of All-Time "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   25.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   960
      TabIndex        =   10
      Top             =   480
      Width           =   12732
   End
   Begin VB.Label lblSubtitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Enhace Your Video Game Knowledge!   How Many of the Top 20 Games are Mario Related?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3960
      TabIndex        =   9
      Top             =   1560
      Width           =   9012
   End
   Begin VB.Label lblInformation 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This information was collected in 2002"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3840
      TabIndex        =   8
      Top             =   7680
      Width           =   6972
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Mario Madness
'Form name: frmSearch
'Author: Bill Macy
'Date Written: Tuesday March 14th, 2006
'Objective of form:  This form allows the user to look through a bunch of information on the top 20 video games of all times
                'in 2002.  They can sort it according to a variety of titles.  They first load the data and view it, then
                'they can sort is according to the amount of games sold, the system it was on, the month it was released,
                'the rank it was in, and the name of the game.  All they have to do is click on the button they want to
                'arrange the data by.  They can also return to the main page.



Option Explicit
Public counter As Integer       'declares all my variables
Public i As Integer
Dim Rank(1 To 21) As Integer
Dim Game(1 To 21) As String
Dim System(1 To 21) As String
Dim Numbersold(1 To 21) As String
Dim Release(1 To 21) As String
Public pos, pass As Integer
Public temprank As Integer
Public tempgame As String
Public tempsystem As String
Public tempnumbersold As String
Public temprelease As String
Public size As Integer

Private Sub CmdGamesSold_Click()
    For pass = 1 To (size - 1)      'loops through to sort by the number of games sold
        For counter = 1 To (size - pass)
            If Rank(counter) > Rank(counter + 1) Then        'checks to see if the current rank is bigger then the next rank and loops through if it is
                tempnumbersold = Numbersold(counter)        'performs the switch by taking the bigger and storing it in a temp variable
                Numbersold(counter) = Numbersold(counter + 1)      'moves the smaller one where the bigger used to be
                Numbersold(counter + 1) = tempnumbersold          'places the bigger after the smaller so they are in order
                
                tempgame = Game(counter)        'see above
                Game(counter) = Game(counter + 1)
                Game(counter + 1) = tempgame
                
                temprank = Rank(counter)        'see above
                Rank(counter) = Rank(counter + 1)
                Rank(counter + 1) = temprank
                
                tempsystem = System(counter)        'see above
                System(counter) = System(counter + 1)
                System(counter + 1) = tempsystem
                
                temprelease = Release(counter)      'see above
                Release(counter) = Release(counter + 1)
                Release(counter + 1) = temprelease
            End If
        Next counter
    Next pass
    picresults.Cls      'clears the picture box of anything previously written
    picresults.Print "Rank"; Tab(10); "Game Title"; Tab(50); "System", "Number of Games Sold (In Millions)", "Release Date"     'prints the titles
    picresults.Print "*****************************************************************************************************************************************************************************"        'prints a bunch of stars
    For i = 1 To size       'goes through and prints all the data in the correct order
        picresults.Print Rank(i); Tab(10); Game(i); Tab(50); System(i), , Numbersold(i), , , Release(i)     'prints the Data aligned
    Next i      'loops
End Sub

Private Sub cmdreturn_Click()
    frmSearch.Hide      'hides the search page
    frmMain.Show        'shows the main page
End Sub


Private Sub cmdLoad_Click()
    picresults.Cls      'clears out the picture box of anything that may have been in it previously
    counter = 0
    Open App.Path & "\TopList.txt" For Input As #3      'opens the file associated with the data
    Do Until EOF(3)     'read the entire file
        counter = counter + 1       'incremented by one so the array is filled appropriatly
        Input #3, Rank(counter), Game(counter), System(counter), Numbersold(counter), Release(counter)      'fills the array with the data
    Loop        'loops bacak around
    Close #3        'closes the file
    size = counter      'sets a variable equal to the size of the file
    picresults.Print "Rank"; Tab(10); "Game Title"; Tab(50); "System", "Number of Games Sold (In Millions)", "Release Date"     'prints the titles
    picresults.Print "*****************************************************************************************************************************************************************************"        'prints a bunch of stars
    For i = 1 To counter        'loops throught to print all of the data in the correct order
        picresults.Print Rank(i); Tab(10); Game(i); Tab(50); System(i), , Numbersold(i), , , Release(i)     'prints the data
    Next i      'loops throught the array
End Sub

Private Sub cmdRank_Click()

    For pass = 1 To (size - 1)      'see above example
        For counter = 1 To (size - pass)
            If Rank(counter) > Rank(counter + 1) Then
                tempnumbersold = Numbersold(counter)
                Numbersold(counter) = Numbersold(counter + 1)
                Numbersold(counter + 1) = tempnumbersold
                
                tempgame = Game(counter)
                Game(counter) = Game(counter + 1)
                Game(counter + 1) = tempgame
                
                temprank = Rank(counter)
                Rank(counter) = Rank(counter + 1)
                Rank(counter + 1) = temprank
                
                tempsystem = System(counter)
                System(counter) = System(counter + 1)
                System(counter + 1) = tempsystem
                
                temprelease = Release(counter)
                Release(counter) = Release(counter + 1)
                Release(counter + 1) = temprelease
            End If
        Next counter
    Next pass
    picresults.Cls
    picresults.Print "Rank"; Tab(10); "Game Title"; Tab(50); "System", "Number of Games Sold (In Millions)", "Release Date"
    picresults.Print "*****************************************************************************************************************************************************************************"
    For i = 1 To size
        picresults.Print Rank(i); Tab(10); Game(i); Tab(50); System(i), , Numbersold(i), , , Release(i)
    Next i
End Sub

Private Sub cmdSystem_Click()
    
    For pass = 1 To (size - 1)      'see above example
        For counter = 1 To (size - pass)
            If System(counter) > System(counter + 1) Then
                tempnumbersold = Numbersold(counter)
                Numbersold(counter) = Numbersold(counter + 1)
                Numbersold(counter + 1) = tempnumbersold
                
                tempgame = Game(counter)
                Game(counter) = Game(counter + 1)
                Game(counter + 1) = tempgame
                
                temprank = Rank(counter)
                Rank(counter) = Rank(counter + 1)
                Rank(counter + 1) = temprank
                
                tempsystem = System(counter)
                System(counter) = System(counter + 1)
                System(counter + 1) = tempsystem
                
                temprelease = Release(counter)
                Release(counter) = Release(counter + 1)
                Release(counter + 1) = temprelease
            End If
        Next counter
    Next pass
    picresults.Cls
    picresults.Print "Rank"; Tab(10); "Game Title"; Tab(50); "System", "Number of Games Sold (In Millions)", "Release Date"
    picresults.Print "*****************************************************************************************************************************************************************************"
    For i = 1 To size
        picresults.Print Rank(i); Tab(10); Game(i); Tab(50); System(i), , Numbersold(i), , , Release(i)
    Next i
End Sub
Private Sub CmdGame_Click()
    For pass = 1 To (size - 1)      'see above example
        For counter = 1 To (size - pass)
            If Game(counter) > Game(counter + 1) Then
                tempnumbersold = Numbersold(counter)
                Numbersold(counter) = Numbersold(counter + 1)
                Numbersold(counter + 1) = tempnumbersold
                
                tempgame = Game(counter)
                Game(counter) = Game(counter + 1)
                Game(counter + 1) = tempgame
                
                temprank = Rank(counter)
                Rank(counter) = Rank(counter + 1)
                Rank(counter + 1) = temprank
                
                tempsystem = System(counter)
                System(counter) = System(counter + 1)
                System(counter + 1) = tempsystem
                
                temprelease = Release(counter)
                Release(counter) = Release(counter + 1)
                Release(counter + 1) = temprelease
            End If
        Next counter
    Next pass
    picresults.Cls
    picresults.Print "Rank"; Tab(10); "Game Title"; Tab(50); "System", "Number of Games Sold (In Millions)", "Release Date"
    picresults.Print "*****************************************************************************************************************************************************************************"
    For i = 1 To size
        picresults.Print Rank(i); Tab(10); Game(i); Tab(50); System(i), , Numbersold(i), , , Release(i)
    Next i
End Sub

Private Sub CmdRelease_Click()
    For pass = 1 To (size - 1)      'see above example
        For counter = 1 To (size - pass)
            If Release(counter) > Release(counter + 1) Then
                tempnumbersold = Numbersold(counter)
                Numbersold(counter) = Numbersold(counter + 1)
                Numbersold(counter + 1) = tempnumbersold
                
                tempgame = Game(counter)
                Game(counter) = Game(counter + 1)
                Game(counter + 1) = tempgame
                
                temprank = Rank(counter)
                Rank(counter) = Rank(counter + 1)
                Rank(counter + 1) = temprank
                
                tempsystem = System(counter)
                System(counter) = System(counter + 1)
                System(counter + 1) = tempsystem
                
                temprelease = Release(counter)
                Release(counter) = Release(counter + 1)
                Release(counter + 1) = temprelease
            End If
        Next counter
    Next pass
    picresults.Cls
    picresults.Print "Rank"; Tab(10); "Game Title"; Tab(50); "System", "Number of Games Sold (In Millions)", "Release Date"
    picresults.Print "*****************************************************************************************************************************************************************************"
    For i = 1 To size
        picresults.Print Rank(i); Tab(10); Game(i); Tab(50); System(i), , Numbersold(i), , , Release(i)
    Next i
End Sub
