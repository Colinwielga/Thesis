VERSION 5.00
Begin VB.Form frmRockandRoll 
   BackColor       =   &H00400040&
   Caption         =   "Rock and Roll"
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdditon 
      Caption         =   "How many years after the first band started did the last band form? "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   7
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   6
      Top             =   4680
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   10095
      Left            =   120
      ScaleHeight     =   10035
      ScaleWidth      =   5475
      TabIndex        =   5
      Top             =   120
      Width           =   5535
   End
   Begin VB.CommandButton cmdYearSort 
      Caption         =   "Sort all the bands by year the band began. "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   4
      Top             =   2400
      Width           =   2775
   End
   Begin VB.CommandButton cmdAlphabet 
      Caption         =   "Sort all bands alphabetically."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "See list of bands that are classified as Rock and Roll."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5880
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   1
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   0
      Top             =   5520
      Width           =   1575
   End
End
Attribute VB_Name = "frmRockandRoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Music VB Project by Cassiann Procenko
'Form Name is frmRockandRoll
'Date written 10/16/2009
'Purpose of this form is to load the rock and roll text document and sort it by year and by band name.  It also shows the distance in years between when the first band formed and the last band formed.

Private Sub cmdAdditon_Click()
'define variables
Dim rockbandList(1 To 50) As String, yearList(1 To 50) As Single
Dim CTR As Integer, Pass As Integer, Pos As Integer
Dim Tempyear As Single, Temprockband As String
Const K As Single = 1
Const J As Single = 37
Dim YearsBetween As Single

'print header information
picResults.Print
picResults.Print "Number of Years Between First and Last Bands"
picResults.Print "************************************************************"

'open the file to be read and made into arrays
Open App.Path & "\RockandRoll.txt" For Input As #1

'create arrays
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, rockbandList(CTR), yearList(CTR)
Loop

'sort the file
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If yearList(Pos) < yearList(Pos + 1) Then
            Tempyear = yearList(Pos)
            yearList(Pos) = yearList(Pos + 1)
            yearList(Pos + 1) = Tempyear
                Temprockband = rockbandList(Pos)
                rockbandList(Pos) = rockbandList(Pos + 1)
                rockbandList(Pos + 1) = Temprockband
        End If
    Next Pos
Next Pass

'subtraction formula
YearsBetween = yearList(K) - yearList(J)
'print results
picResults.Print YearsBetween

Close #1
End Sub

Private Sub cmdClear_Click()
'clear the picture box
picResults.Cls
End Sub

Private Sub cmdQuit_Click()
'show and hide forms
    frmLeave.Show
    frmRockandRoll.Hide
End Sub

Private Sub cmdReturn_Click()
'show and hide forms
    frmMusicTypes.Show
    frmRockandRoll.Hide
End Sub

Private Sub cmdSort_Click()

'clear picture box
picResults.Cls

'define variables
Dim rockbandList(1 To 50) As String, yearList(1 To 50) As Single
Dim CTR As Integer

'print header information
picResults.Print
picResults.Print "Band Names"; Tab(30); "Year Formed"
picResults.Print "*********************************************************************"

'open the file to be read and made into arrays
Open App.Path & "\RockandRoll.txt" For Input As #1

'create arrays
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, rockbandList(CTR), yearList(CTR)
    picResults.Print rockbandList(CTR); Tab(30); yearList(CTR)
Loop

Close #1
End Sub

Private Sub cmdAlphabet_Click()

'clear picture box
picResults.Cls

'define variables
Dim CTR As Integer, Pass As Integer, Pos As Integer
Dim rockbandList(1 To 50) As String, yearList(1 To 50) As Single
Dim Temprockband As String, Tempyear As Single, K As Single

'open the file to be read and made into arrays
Open App.Path & "\RockandRoll.txt" For Input As #1

'create arrays
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, rockbandList(CTR), yearList(CTR)
Loop

'Sort the file
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If rockbandList(Pos) > rockbandList(Pos + 1) Then
            Temprockband = rockbandList(Pos)
            rockbandList(Pos) = rockbandList(Pos + 1)
            rockbandList(Pos + 1) = Temprockband
                Tempyear = yearList(Pos)
                yearList(Pos) = yearList(Pos + 1)
                yearList(Pos + 1) = Tempyear
        End If
    Next Pos
Next Pass

'print header
picResults.Print
picResults.Print "Band Names"; Tab(30); "Year Formed"
picResults.Print "************************************************************"

'print results
For K = 1 To CTR
    picResults.Print rockbandList(K); Tab(30); yearList(K)
Next K

Close #1
End Sub

Private Sub cmdYearSort_Click()

'clear picture box
picResults.Cls

'define variables
Dim CTR As Integer, Pass As Single, Pos As Single
Dim rockbandList(1 To 50) As String, yearList(1 To 50) As Single
Dim Temprockband As String, Tempyear As Single, J As Integer

'open the file to be read and made into arrays
Open App.Path & "\RockandRoll.txt" For Input As #1

CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, rockbandList(CTR), yearList(CTR)
Loop

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If yearList(Pos) < yearList(Pos + 1) Then
            Tempyear = yearList(Pos)
            yearList(Pos) = yearList(Pos + 1)
            yearList(Pos + 1) = Tempyear
                Temprockband = rockbandList(Pos)
                rockbandList(Pos) = rockbandList(Pos + 1)
                rockbandList(Pos + 1) = Temprockband
        End If
    Next Pos
Next Pass

'print header
picResults.Print
picResults.Print "Band Names"; Tab(30); "Year Formed"
picResults.Print "******************************************************************"

'print results
For J = 1 To CTR
    picResults.Print rockbandList(J); Tab(30); yearList(J)
Next J
Close #1
End Sub
