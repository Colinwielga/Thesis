VERSION 5.00
Begin VB.Form frmPlayerInfo 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   12330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18825
   FillColor       =   &H8000000E&
   BeginProperty Font 
      Name            =   "Lucida Bright"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   Picture         =   "frmPlayerInfo.frx":0000
   ScaleHeight     =   12330
   ScaleWidth      =   18825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to the Start Form"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16440
      TabIndex        =   8
      Top             =   10080
      Width           =   1695
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Information"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      TabIndex        =   7
      Top             =   6480
      Width           =   2055
   End
   Begin VB.PictureBox picPlayerStats 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   12000
      ScaleHeight     =   1755
      ScaleWidth      =   3795
      TabIndex        =   5
      Top             =   7320
      Width           =   3855
   End
   Begin VB.TextBox txtLastName 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   4
      Top             =   5640
      Width           =   5175
   End
   Begin VB.PictureBox picPicture 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   9840
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmdBattingAverage 
      Caption         =   "Calculate the Current Team Batting Average"
      Height          =   1575
      Left            =   10320
      TabIndex        =   2
      Top             =   960
      Width           =   5775
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00800000&
      Caption         =   "Sort and Display Current Players In Order of their Jersey Numbers"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin VB.PictureBox picSortResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   10815
      Left            =   720
      ScaleHeight     =   10755
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   1200
      Width           =   6975
   End
   Begin VB.Label lblPlayerCommand 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enter the last name of a current Minnesota Twins player to view both a picture and statistics about that player:"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   9720
      TabIndex        =   6
      Top             =   4800
      Width           =   6495
   End
End
Attribute VB_Name = "frmPlayerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form deals with player information

Private Sub cmdBattingAverage_Click() 'sum batting average for the players and calculte the average, excluding those players with 0 as their batting average
Dim A As Integer, Count As Integer, Average As Single, TotalAverage As Single, Found As Boolean

Count = 0

For A = 1 To Ctr
    If PlayerBattingAvg(A) > 0 Then 'execute this loop only if the player has a batting average (so not for pitchers)
        Count = Count + 1           'add one to the count
        Average = Average + PlayerBattingAvg(A) 'add the players individual average to the cumulative average
    End If
Next A      'repeat for each player

TotalAverage = Average / Count 'calculate total average by dividing the sum of the player averages by the number of players

MsgBox "The combined Minnesota Twins batting average for 2010 was " & FormatNumber(TotalAverage, 3) & "." 'print results in a messagebox

End Sub

Private Sub cmdDisplay_Click()  'After the user enters the name of a player, display that player's picture and information about them
Dim InputName As String, M As Integer

picPlayerStats.Cls  'clear the picture boxes corresponding to this action to prepare for new input
picPicture = LoadPicture()

'Get Text from inputbox
InputName = txtLastName.Text

For M = 1 To Ctr    'read through array to find name entered in textbox
    If InputName = LastName(M) Then
        picPlayerStats.Print "First Name:"; Tab(20); FirstName(M)   'print the statistics corresponding to the player
        picPlayerStats.Print "Last Name:"; Tab(20); LastName(M)
        picPlayerStats.Print "Jersey Number:"; Tab(20); PlayerNumber(M)
        picPlayerStats.Print "Position:"; Tab(20); Position(M)
            If PlayerBattingAvg(M) > 0 Then     'only print the batting average if the player has one
                picPlayerStats.Print "Batting Average:"; Tab(20); PlayerBattingAvg(M)
            End If
        picPlayerStats.Print "Birthdate:"; Tab(20); Birthdate(M)
    End If
Next M

Select Case InputName   'input a picture of the player whose name was entered in the textbox
Case Is = "Span"
    picPicture = LoadPicture("M:\CS130\VBProject\Span.jpg")
Case Is = "Young"
    picPicture = LoadPicture("M:\CS130\VBProject\Young.jpg")
Case Is = "Cuddyer"
    picPicture = LoadPicture("M:\CS130\VBProject\Cuddyer.jpg")
Case Is = "Thome"
    picPicture = LoadPicture("M:\CS130\VBProject\Thome.jpg")
Case Is = "Kubel"
    picPicture = LoadPicture("M:\CS130\VBProject\Kubel.jpg")
Case Is = "Revere"
    picPicture = LoadPicture("M:\CS130\VBProject\Revere.jpg")
Case Is = "Repko"
    picPicture = LoadPicture("M:\CS130\VBProject\Repko.jpg")
Case Is = "Baker"
    picPicture = LoadPicture("M:\CS130\VBProject\Baker.jpg")
Case Is = "Blackburn"
    picPicture = LoadPicture("M:\CS130\VBProject\Blackburn.jpg")
Case Is = "Burnett"
    picPicture = LoadPicture("M:\CS130\VBProject\Burnett.jpg")
Case Is = "Capps"
    picPicture = LoadPicture("M:\CS130\VBProject\Capps.jpg")
Case Is = "Crain"
    picPicture = LoadPicture("M:\CS130\VBProject\Crain.jpg")
Case Is = "Duensing"
    picPicture = LoadPicture("M:\CS130\VBProject\Duensing.jpg")
Case Is = "Flores"
    picPicture = LoadPicture("M:\CS130\VBProject\Flores.jpg")
Case Is = "Condrey"
    picPicture = LoadPicture("M:\CS130\VBProject\Condrey.jpg")
Case Is = "Delaney"
    picPicture = LoadPicture("M:\CS130\VBProject\Delaney.jpg")
Case Is = "Guerra"
    picPicture = LoadPicture("M:\CS130\VBProject\Guerra.jpg")
Case Is = "Guerrier"
    picPicture = LoadPicture("M:\CS130\VBProject\Guerrier.jpg")
Case Is = "Slama"
    picPicture = LoadPicture("M:\CS130\VBProject\Slama.jpg")
Case Is = "Fuentes"
    picPicture = LoadPicture("M:\CS130\VBProject\Fuentes.jpg")
Case Is = "Liriano"
    picPicture = LoadPicture("M:\CS130\VBProject\Liriano.jpg")
Case Is = "Mahay"
    picPicture = LoadPicture("M:\CS130\VBProject\Mahay.jpg")
Case Is = "Manship"
    picPicture = LoadPicture("M:\CS130\VBProject\Manship.jpg")
Case Is = "Mijares"
    picPicture = LoadPicture("M:\CS130\VBProject\Mijares.jpg")
Case Is = "Nathan"
    picPicture = LoadPicture("M:\CS130\VBProject\Nathan.jpg")
Case Is = "Neshek"
    picPicture = LoadPicture("M:\CS130\VBProject\Neshek.jpg")
Case Is = "Pavano"
    picPicture = LoadPicture("M:\CS130\VBProject\Pavano.jpg")
Case Is = "Perkins"
    picPicture = LoadPicture("M:\CS130\VBProject\Perkins.jpg")
Case Is = "Rauch"
    picPicture = LoadPicture("M:\CS130\VBProject\Rauch.jpg")
Case Is = "Slowey"
    picPicture = LoadPicture("M:\CS130\VBProject\Slowey.jpg")
Case Is = "Swarzak"
    picPicture = LoadPicture("M:\CS130\VBProject\Swarzak.jpg")
Case Is = "Butera"
    picPicture = LoadPicture("M:\CS130\VBProject\Butera.jpg")
Case Is = "Mauer"
    picPicture = LoadPicture("M:\CS130\VBProject\Mauer.jpg")
Case Is = "Morales"
    picPicture = LoadPicture("M:\CS130\VBProject\Morales.jpg")
Case Is = "Casilla"
    picPicture = LoadPicture("M:\CS130\VBProject\Casilla.jpg")
Case Is = "Hardy"
    picPicture = LoadPicture("M:\CS130\VBProject\Hardy.jpg")
Case Is = "Hudson"
    picPicture = LoadPicture("M:\CS130\VBProject\Hudson.jpg")
Case Is = "Hughes"
    picPicture = LoadPicture("M:\CS130\VBProject\Hughes.jpg")
Case Is = "Morneau"
    picPicture = LoadPicture("M:\CS130\VBProject\Morneau.jpg")
Case Is = "Plouffe"
    picPicture = LoadPicture("M:\CS130\VBProject\Plouffe.jpg")
Case Is = "Punto"
    picPicture = LoadPicture("M:\CS130\VBProject\Punto.jpg")
Case Is = "Tolbert"
    picPicture = LoadPicture("M:\CS130\VBProject\Tolbert.jpg")
Case Is = "Valencia"
    picPicture = LoadPicture("M:\CS130\VBProject\Valencia.jpg")
Case Else   'If the name entered does not match, clear the textbox and display an error message
    txtLastName.Text = ""
    MsgBox "Error. Please enter a valid name. (Note: The first letter must be capitalized.)"
End Select

txtLastName.Text = ""   'clear the textbox when the picture and data are displayed, allowing user to enter a new name


End Sub

Private Sub cmdReturn_Click() 'Transfer between forms
    frmStart.Show   'return to the start form
    frmPlayerInfo.Hide
    
    picPicture = LoadPicture() 'clear box of player's picture if user returns to start form
End Sub

Private Sub cmdSort_Click() 'sort players by their number
Dim Pos As Integer, Pass As Integer, K As Integer
Dim TempLastName As String, TempFirstName As String, TempPlayerNumber As Long
Dim TempPosition As String, TempPlayerBattingAvg As String, TempBirthdate As String

picSortResults.Cls  'clear picturebox

For Pass = 1 To Ctr - 1         'keep track of how many passes
    For Pos = 1 To Ctr - Pass 'keep track of how many comparisons
        If PlayerNumber(Pos) > PlayerNumber(Pos + 1) Then
            TempPlayerNumber = PlayerNumber(Pos) 'exchange player numbers
            PlayerNumber(Pos) = PlayerNumber(Pos + 1)
            PlayerNumber(Pos + 1) = TempPlayerNumber
            
            'exchange corresponding information for each player
            TempLastName = LastName(Pos) 'exchange last name if out of order
            LastName(Pos) = LastName(Pos + 1)
            LastName(Pos + 1) = TempLastName
            
            TempFirstName = FirstName(Pos) 'exchange first names
            FirstName(Pos) = FirstName(Pos + 1)
            FirstName(Pos + 1) = TempFirstName
            
            TempPosition = Position(Pos) 'exchange player's positions
            Position(Pos) = Position(Pos + 1)
            Position(Pos + 1) = TempPosition
            
            TempPlayerBattingAvg = PlayerBattingAvg(Pos) 'exchange player's batting averages
            PlayerBattingAvg(Pos) = PlayerBattingAvg(Pos + 1)
            PlayerBattingAvg(Pos + 1) = TempPlayerBattingAvg
            
            TempBirthdate = Birthdate(Pos) 'exchange player's birthdates
            Birthdate(Pos) = Birthdate(Pos + 1)
            Birthdate(Pos + 1) = TempBirthdate
        End If
    Next Pos
Next Pass

'print heading for table
picSortResults.Print "Current Minnesota Twins Players:"
picSortResults.Print
picSortResults.Print "Jersey Number"; Tab(20); "First Name"; Tab(40); "Last Name"
picSortResults.Print "***********************************************************************"

'print the sorted list
For K = 1 To Ctr
    picSortResults.Print Tab(6); PlayerNumber(K); Tab(20); FirstName(K); Tab(40); LastName(K)
Next K



End Sub
