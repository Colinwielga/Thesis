VERSION 5.00
Begin VB.Form frmPokemonSort 
   Caption         =   "Pokemon Sort"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form5"
   ScaleHeight     =   4950
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   615
      Left            =   5160
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdType1 
      Caption         =   "Sort By Type"
      Height          =   615
      Left            =   4200
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdName 
      Caption         =   "Sort By Name"
      Height          =   615
      Left            =   3360
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdOutput 
      Caption         =   "Output Current Data"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Menu"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "Sort by Pokedex Number"
      Height          =   615
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      Height          =   1455
      Left            =   2040
      Picture         =   "Pokemon Sort.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
   End
   Begin VB.PictureBox Picture3 
      Height          =   1095
      Left            =   3840
      Picture         =   "Pokemon Sort.frx":0C95
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   3480
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   120
      Picture         =   "Pokemon Sort.frx":1A3A
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Current Data"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox pbxResults 
      Height          =   2295
      Left            =   1200
      ScaleHeight     =   2235
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   840
      Width           =   4695
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Pokemon Data"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Designed by Chris Davin"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   1935
   End
End
Attribute VB_Name = "frmPokemonSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : MemoryGamesEtc (Chris Davin's VB Project.vbp)
'Form Name : frmPokemonSort (Pokemon Sort.frm)
'Author: Chris Davin
'Date Written: October 29, 2003
'Purpose of Form: To take data from an array outside which could be increased
                 'with additional data entered.  Sorts the info. accourding to
                 'the three parts of each entry.  Can output data then so
                 'user will have useful sorted Pojkemon data.  Can also search
                 'for a specific entry.

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Dim PokeData As String
Dim NewPokeData As String
Dim Pokemon(1 To 150) As String, Type1(1 To 150) As String, temp As String
Dim Pokedex(1 To 150) As Integer, icount As Integer, pass As Integer, k As Integer
'A Bubble Sort based on the Pokemon's name
Private Sub cmdName_Click()
    For pass = 1 To icount - 1
            For k = 1 To icount - pass
                If Pokemon(k) > Pokemon(k + 1) Then
                    temp = Pokemon(k)
                    Pokemon(k) = Pokemon(k + 1)
                    Pokemon(k + 1) = temp
                    temp = Type1(k)
                    Type1(k) = Type1(k + 1)
                    Type1(k + 1) = temp
                    temp = Pokedex(k)
                    Pokedex(k) = Pokedex(k + 1)
                    Pokedex(k + 1) = temp
                End If
            Next k
        Next pass
End Sub
'A Bubble Sort based on the Pokemon's Pokedex Number
Private Sub cmdNumber_Click()
    For pass = 1 To icount - 1
        For k = 1 To icount - pass
            If Pokedex(k) > Pokedex(k + 1) Then
                temp = Pokemon(k)
                Pokemon(k) = Pokemon(k + 1)
                Pokemon(k + 1) = temp
                temp = Type1(k)
                Type1(k) = Type1(k + 1)
                Type1(k + 1) = temp
                temp = Pokedex(k)
                Pokedex(k) = Pokedex(k + 1)
                Pokedex(k + 1) = temp
            End If
        Next k
    Next pass
End Sub
'Inputs the data into an array
Private Sub cmdLoad_Click()
    icount = 0
    Open App.Path & "\j.t.txt" For Input As #1
        Do While Not EOF(1)
            k = k + 1
            Input #1, Pokemon(k), Type1(k), Pokedex(k)
            icount = icount + 1
        Loop
    Close #1
End Sub
'Outputs data into a new file
Private Sub cmdOutput_Click()
    Open App.Path & "\NewPokeData" For Output As #2
        For k = 1 To icount
            Print #2, Pokemon(k), Type1(k), Pokedex(k)
        Next k
    Close #2
End Sub
'Shows how current data looks
Private Sub cmdPrint_Click()
    pbxResults.Cls
    For k = 1 To icount
        pbxResults.Print Pokemon(k), Type1(k), Pokedex(k)
    Next k
End Sub
'Quits program
Private Sub cmdQuit_Click()
    End
End Sub
'Returns to main menu
Private Sub cmdReturn_Click()
    frmPokemonSort.Hide
    frmMainMenu.Show
End Sub
'A search for specific Pokemon name
'Useful if more than elevin entries
Private Sub cmdSearch_Click()
    Dim NotFound As Boolean
    pbxResults.Cls
    Dim Found As Boolean
    Dim N As String
    N = InputBox("Enter the name of the pokemon to search for.", "Info.")
    k = 0
    NotFound = True
    If N = "" Then NotFound = False
    Do While NotFound
        k = k + 1
        If k > icount Then
            pbxResults.Print N; " is not in the database."
            NotFound = False
        End If
        If N = Pokemon(k) Then
            NotFound = False
            pbxResults.Print N; " is a "; Type1(k); " type, with a Pokedex number of "; Pokedex(k)
        End If
    Loop
End Sub
'A Bubble Sort by Pokemon Type
Private Sub cmdType1_Click()
    For pass = 1 To icount - 1
            For k = 1 To icount - pass
                If Type1(k) > Type1(k + 1) Then
                    temp = Pokemon(k)
                    Pokemon(k) = Pokemon(k + 1)
                    Pokemon(k + 1) = temp
                    temp = Type1(k)
                    Type1(k) = Type1(k + 1)
                    Type1(k + 1) = temp
                    temp = Pokedex(k)
                    Pokedex(k) = Pokedex(k + 1)
                    Pokedex(k + 1) = temp
                End If
            Next k
        Next pass
End Sub

Private Sub Form_Load()
    Path = "N:\CS130\handin\Chris Davin's VBProject\"
    PokeData = Path & "Data Files\Pokemon.txt"
    NewPokeData = Path & "Data Files\NewPokeData.txt"
End Sub

'Info. on the picture
Private Sub Picture1_Click()
    MsgBox "This is Vulpix.", , "Info."
End Sub
'Info. on the picture
Private Sub Picture3_Click()
    MsgBox "These are baby Pikachus called Pichus.", , "Info."
End Sub
'Info. on the picture
Private Sub Picture4_Click()
    MsgBox "This is Pikachu, he is yellow.", , "Info."
End Sub
