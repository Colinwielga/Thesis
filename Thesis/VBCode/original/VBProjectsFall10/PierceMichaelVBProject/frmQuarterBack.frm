VERSION 5.00
Begin VB.Form frmQuarterBack 
   BackColor       =   &H80000006&
   Caption         =   "Form1"
   ClientHeight    =   9300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   15195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnReadQuarterBackData 
      BackColor       =   &H80000005&
      Caption         =   "1. Load Quarter Back Statistics"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton btnShowTop10QuarterBacks 
      Caption         =   "View Top 10 Quarter Backs by Overall Rank"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   3255
   End
   Begin VB.CommandButton btnSortByYardsQB 
      Caption         =   "Sort Quarter Backs by Their Amount of Yards Gained"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton btnSortByTdsQB 
      Caption         =   "Sort Quarter Backs by Their Number of Touch Downs Scored"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton btnQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   4
      Top             =   7800
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   3840
      ScaleHeight     =   3795
      ScaleWidth      =   10755
      TabIndex        =   3
      Top             =   3000
      Width           =   10815
   End
   Begin VB.CommandButton btnReturnToHome 
      Caption         =   "Return To Home Screen"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   6480
      Width           =   3255
   End
   Begin VB.CommandButton btnSearchYardsQB 
      Caption         =   "Find Players by Amount of Yards"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton btnSearchTdsQB 
      Caption         =   "Find Players by Number of Touchdowns"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   1995
      Left            =   11760
      Picture         =   "frmQuarterBack.frx":0000
      Top             =   960
      Width           =   2610
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "2009 NFL Quarter Back Statistics"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   14895
   End
End
Attribute VB_Name = "frmQuarterBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare variables for entire form
Dim player(1 To 10) As String
Dim team(1 To 10) As String
Dim yards(1 To 10) As Single
Dim tds(1 To 10) As Single
Dim ctr As Integer
Option Explicit
Private Sub btnClear_Click() 'this button will clear the picture box when pushed
    picResults.Cls
End Sub
Private Sub btnQuit_Click() 'this button will quit out of program
    End
End Sub

Private Sub btnReadQuarterBackData_Click() 'this button will load the quarterback data from a text file
'open quarter back file
Open App.Path & "\quarterBacks.txt" For Input As #2

ctr = 0

Do While Not EOF(2) 'search through list until complete
    ctr = ctr + 1
    Input #2, player(ctr), team(ctr), yards(ctr), tds(ctr) 'file into 4 parallel arrays
Loop 'loop through file until complete

'alert user that the list was loaded
MsgBox "The list of quarter backs was successfully loaded."

End Sub
Private Sub btnReturnToHome_Click() 'this button will hide the current form and return user to home screen
    frmStart.Show
    frmQuarterBack.Hide
End Sub
Private Sub btnSearchTdsQB_Click() 'this button will find all the players who have scored more touchdowns than the amount entered by user
'Dim variables
Dim Found As Boolean
Dim I As Integer
Dim tdsWanted As String

picResults.Cls 'clear picture box for new results

'assign variables
'user will enter amount of touchdowns into an input box
tdsWanted = InputBox("Enter the number of touchdowns that you wish to see all of the quarter backs who have thrown for more touchdowns than the amount entered.")
Found = False

picResults.Print "The quarter backs who have thrown for more than "; tdsWanted; " touchdowns are:"
picResults.Print
picResults.Print "Name"; Tab(25); "Team", "Touchdowns" 'set up table headings
picResults.Print "*********************************************************************************************"

For I = 1 To ctr 'searches the entire list
    If tdsWanted < tds(I) Then
        Found = True
        picResults.Print player(I); Tab(25); team(I), , tds(I)
    End If
Next I

If (Not Found) Then 'if amount is not found this message will be displayed
    MsgBox "No player has thrown for more touchdowns than " & tdsWanted & " touchdowns are:"
End If

End Sub

Private Sub btnSearchYardsQB_Click() 'this button will find all players who have thrown for more than the entered amount of yards
'Dim variables
Dim Found As Boolean
Dim I As Integer
Dim yardsWanted As String

picResults.Cls 'clear picture box for new results

'assign variables
'user will enter amount of yards they want to search for in an input box
yardsWanted = InputBox("Enter an amount of yards that you wish to see all of the quarter backs who have thrown for more than the amount entered.")
Found = False

picResults.Print "The quarter backs who have thrown for more than "; yardsWanted; " yards are:"
picResults.Print
picResults.Print "Name"; Tab(25); "Team", "Yards"
picResults.Print "********************************************************************************************"

For I = 1 To ctr 'searches the entire list
    If yardsWanted < yards(I) Then
        Found = True
        picResults.Print player(I); Tab(25); team(I), , yards(I)
    End If
Next I

If (Not Found) Then 'if amount is not found this message will be displayed
    MsgBox "No quarter back has thrown for more than " & yardsWanted & " yards are:"
End If

End Sub

Private Sub btnShowTop10QuarterBacks_Click() 'this button will show the top 10 ranked quarterbacks from original list
Dim J As Integer

picResults.Cls 'clear picture box for new results

J = 0

picResults.Print "Name"; Tab(25); "Team", "Yards", "Touchdowns" 'headings for table
picResults.Print "*************************************************************************"
For J = 1 To ctr 'prints original list
    picResults.Print J; player(J); Tab(25); team(J), , yards(J), tds(J)
Next J 'goes to next player

End Sub

Private Sub btnSortByTdsQB_Click() 'this button will sort data by the number of touchdowns in descending order
Dim TempName As String
Dim TempTeam As String
Dim TempTds As Single
Dim Pass As Integer
Dim I As Integer
Dim Pos As Integer

picResults.Cls 'clear picture box for new results

For Pass = 1 To ctr - 1 'begin bubble search to sort players
    For Pos = 1 To ctr - Pass 'keep track of how many comparasions
        If tds(Pos) < tds(Pos + 1) Then
            TempTds = tds(Pos) 'swaps values if out of order
            tds(Pos) = tds(Pos + 1)
            tds(Pos + 1) = TempTds
            
            TempName = player(Pos) 'swaps values if out of order
            player(Pos) = player(Pos + 1)
            player(Pos + 1) = TempName
            
            TempTeam = team(Pos) 'swaps values if out of order
            team(Pos) = team(Pos + 1)
            team(Pos + 1) = TempTeam
        End If
    Next Pos 'move to next position in list
Next Pass 'do another pass

picResults.Print "Quarter Backs sorted by the number of touchdowns earned:"
picResults.Print
picResults.Print "Player"; Tab(25); "Team", "Touchdowns"
picResults.Print "*********************************************************************"
For I = 1 To ctr
    picResults.Print I; player(I); Tab(25); team(I), , tds(I)
Next I

End Sub

Private Sub btnSortByYardsQB_Click() 'this button re orders data in descending order by amount of yards thrown
Dim TempName As String
Dim TempTeam As String
Dim TempYards As Single
Dim Pass As Integer
Dim I As Integer
Dim Pos As Integer

picResults.Cls 'clear picture box for new results

For Pass = 1 To ctr - 1 'begin bubble search to sort players
    For Pos = 1 To ctr - Pass 'keep track of how many comparasions
        If yards(Pos) < yards(Pos + 1) Then
            TempYards = yards(Pos) 'swaps values if out of order
            yards(Pos) = yards(Pos + 1)
            yards(Pos + 1) = TempYards
            
            TempName = player(Pos) 'swaps values if out of order
            player(Pos) = player(Pos + 1)
            player(Pos + 1) = TempName
            
            TempTeam = team(Pos) 'swaps values if out of order
            team(Pos) = team(Pos + 1)
            team(Pos + 1) = TempTeam
        End If
    Next Pos 'move to next position in list
Next Pass 'do another pass

picResults.Print "Quarter backs sorted by amount of yards gained:"
picResults.Print
picResults.Print "Player"; Tab(25); "Team", "Yards" 'headings for table
picResults.Print "*********************************************************************"
For I = 1 To ctr
    picResults.Print I; player(I); Tab(25); team(I), , yards(I)
Next I

End Sub



