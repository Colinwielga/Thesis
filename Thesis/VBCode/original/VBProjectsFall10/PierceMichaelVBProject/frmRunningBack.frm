VERSION 5.00
Begin VB.Form frmRunningBack 
   BackColor       =   &H80000006&
   Caption         =   "Form1"
   ClientHeight    =   10050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   10050
   ScaleWidth      =   15090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSearchTds 
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
      Left            =   7920
      TabIndex        =   10
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton btnSearchYards 
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
      Left            =   4080
      TabIndex        =   9
      Top             =   1440
      Width           =   3255
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
      Left            =   360
      TabIndex        =   8
      Top             =   6720
      Width           =   3255
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
      Left            =   4080
      ScaleHeight     =   3795
      ScaleWidth      =   10635
      TabIndex        =   7
      Top             =   3240
      Width           =   10695
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
      Left            =   2160
      TabIndex        =   6
      Top             =   8040
      Width           =   1455
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
      Left            =   360
      TabIndex        =   5
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton btnSortByTds 
      Caption         =   "Sort Running Backs by Their Number of Touch Downs Scored"
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
      Left            =   360
      TabIndex        =   4
      Top             =   5400
      Width           =   3255
   End
   Begin VB.CommandButton btnSortByYards 
      Caption         =   "Sort Running Backs by Their Amount of Yards Gained"
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
      Left            =   360
      TabIndex        =   3
      Top             =   4080
      Width           =   3255
   End
   Begin VB.CommandButton btnShowTop10RunningBacks 
      Caption         =   "View Top 10 Running Backs by Overall Rank"
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
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton btnReadRunningBackData 
      Caption         =   "1. Load Running Back Statistics"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   2925
      Left            =   11760
      Picture         =   "frmRunningBack.frx":0000
      Top             =   7200
      Width           =   2865
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "2009 NFL Running Back Statistics"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   14895
   End
End
Attribute VB_Name = "frmRunningBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declares variables for the entire form
Dim player(1 To 10) As String
Dim team(1 To 10) As String
Dim yards(1 To 10) As Single
Dim tds(1 To 10) As Single
Dim ctr As Integer
Option Explicit
Private Sub btnClear_Click() 'This button will clear the results box when pushed
    picResults.Cls
End Sub

Private Sub btnQuit_Click() 'This button ends the program when pushed
    End
End Sub

Private Sub btnReadRunningBackData_Click() 'this button loads the list of runningbacks
'open runningback file
Open App.Path & "\runningBacks.txt" For Input As #1

ctr = 0 'set counter to zero

Do While Not EOF(1) 'read file until the end
    ctr = ctr + 1
    Input #1, player(ctr), team(ctr), yards(ctr), tds(ctr) 'file into 4 parallel arrays
Loop 'loop through list until complete

'alert user that the list was loaded
MsgBox "The list of running backs was successfully loaded."

End Sub

Private Sub btnReturnToHome_Click() 'This button will hide the current form and return user to home page
    frmStart.Show
    frmRunningBack.Hide
End Sub

Private Sub btnSearchTds_Click() 'This button searches through the list and finds players who scored more tds than amt entered
'Dim variables
Dim Found As Boolean
Dim I As Integer
Dim tdsWanted As Single

picResults.Cls 'clear picture box for new results

'assign variables
'user enters data wanted into an input box
tdsWanted = InputBox("Enter the number of touchdowns that you wish to see all of the running backs who have gained more than the amount entered.")
Found = False

picResults.Print "The running backs that have scored more than "; tdsWanted; " touchdowns are:"
picResults.Print
picResults.Print "Name"; Tab(25); "Team", "Touchdowns" 'set up table
picResults.Print "*********************************************************************************"

For I = 1 To ctr 'searches the entire list
    If tdsWanted < tds(I) Then 'find all the players who have scored more than the inputted amount
        Found = True
        picResults.Print player(I); Tab(25); team(I), , tds(I) 'print results
    End If
Next I

If (Not Found) Then 'if amount is not found this message will be displayed
    MsgBox "No running back has scored more than " & tdsWanted & " touchdowns."
End If

End Sub

Private Sub btnSearchYards_Click() 'this button will find all players who have rushed for more yards than the amount entered in an inputbox
'Dim variables
Dim Found As Boolean
Dim I As Integer
Dim yardsWanted As Single

picResults.Cls 'clears picture box for new results

'assign variables
'user enters amount of yards to search for in an input box
yardsWanted = InputBox("Enter an amount of yards that you wish to see all of the running backs who have gained more than the amount entered.")
Found = False

picResults.Print "The running backs who have gained more than "; yardsWanted; " yards are:" 'set up table
picResults.Print
picResults.Print "Name"; Tab(25); "Team", "Yards"
picResults.Print "********************************************************************************"

For I = 1 To ctr 'searches the entire list
    If yardsWanted < yards(I) Then
        Found = True
        picResults.Print player(I); Tab(25); team(I), , yards(I)
    End If
Next I

If (Not Found) Then 'if amount is not found this message will be displayed
    MsgBox "No running back has rushed for more than " & yardsWanted & " yards."
End If

End Sub

Private Sub btnShowTop10RunningBacks_Click() 'this button will show the top ten runningbacks in the original list
Dim I As Integer

picResults.Cls 'clear picture box for new results

I = 0

picResults.Print "Name"; Tab(25); "Team", "Yards", "Touchdowns" 'headings for table
picResults.Print "***********************************************************************************"
For I = 1 To ctr 'prints original list
    picResults.Print I; player(I); Tab(25); team(I), , yards(I), tds(I)
Next I 'goes to next player
End Sub

Private Sub btnSortByTds_Click() 'this button searches through list and reorders it by amount of touchdowns scored in decsending order
Dim TempName As String
Dim TempTeam As String
Dim TempTds As Single
Dim Pass As Integer
Dim I As Integer
Dim Pos As Integer

picResults.Cls

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

picResults.Print "Running backs sorted by the number of touchdowns earned:"
picResults.Print
picResults.Print "Player"; Tab(25); "Team", "Touchdowns"
picResults.Print "*************************************************************************************"
For I = 1 To ctr
    picResults.Print I; player(I); Tab(25); team(I), , tds(I)
Next I
End Sub

Private Sub btnSortByYards_Click() 'this button searches through list and re orders it by the amount of yards gained in descending order
Dim TempName As String
Dim TempTeam As String
Dim TempYards As Single
Dim Pass As Integer
Dim I As Integer
Dim Pos As Integer

picResults.Cls

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

picResults.Print "Running Backs sorted by amount of yards gained:"
picResults.Print
picResults.Print "Player"; Tab(25); "Team", "Yards"
picResults.Print "**********************************************************************************"
For I = 1 To ctr
    picResults.Print I; player(I); Tab(25); team(I), , yards(I)
Next I
            
End Sub
