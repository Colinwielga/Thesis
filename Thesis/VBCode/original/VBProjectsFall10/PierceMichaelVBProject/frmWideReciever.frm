VERSION 5.00
Begin VB.Form frmWideReciever 
   BackColor       =   &H80000006&
   Caption         =   "Form1"
   ClientHeight    =   10635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17235
   LinkTopic       =   "Form1"
   ScaleHeight     =   10635
   ScaleWidth      =   17235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSearchRecWR 
      Caption         =   "Find Players by Number of Receptions"
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
      Left            =   11280
      TabIndex        =   12
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton btnSearchTdsWR 
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
      Left            =   7560
      TabIndex        =   11
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton btnSearchYardsWR 
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
      TabIndex        =   10
      Top             =   1440
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
      Height          =   5175
      Left            =   3840
      ScaleHeight     =   5115
      ScaleWidth      =   13035
      TabIndex        =   9
      Top             =   3240
      Width           =   13095
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
      Left            =   2040
      TabIndex        =   8
      Top             =   9360
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
      Left            =   240
      TabIndex        =   7
      Top             =   9360
      Width           =   1455
   End
   Begin VB.CommandButton btnReturnToHome 
      Caption         =   "Return to Home Screen"
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
      Left            =   240
      TabIndex        =   6
      Top             =   8040
      Width           =   3255
   End
   Begin VB.CommandButton btnSortByTds 
      Caption         =   "Sort Wide Receivers by Number of Touchdowns Scored"
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
      Left            =   240
      TabIndex        =   5
      Top             =   6720
      Width           =   3255
   End
   Begin VB.CommandButton btnSortByYards 
      Caption         =   "Sort Wide Receivers by Amount of Yards Gained"
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
      Left            =   240
      TabIndex        =   4
      Top             =   5400
      Width           =   3255
   End
   Begin VB.CommandButton btnSortByReceptions 
      Caption         =   "Sort Wide Receivers by Number of Receptions"
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
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   3255
   End
   Begin VB.CommandButton btnViewReceiverStats 
      Caption         =   "View Top 10 Wide Receivers by Overall Rank"
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
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton btnReadWideReceiverData 
      Caption         =   "1. Load Wide Receiver Stats"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   2820
      Left            =   3960
      Picture         =   "frmWideReciever.frx":0000
      Top             =   8520
      Width           =   2700
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "2009 NFL Wide Receiver Statistics"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   14415
   End
End
Attribute VB_Name = "frmWideReciever"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim player(1 To 10) As String
Dim team(1 To 10) As String
Dim rec(1 To 10) As Single
Dim yards(1 To 10) As Single
Dim tds(1 To 10) As Single
Dim ctr As Integer
Option Explicit
Private Sub btnClear_Click()
    picResults.Cls 'This button will clear the results box when pushed
End Sub

Private Sub btnQuit_Click()
    End 'This button will quit the program when pushed
End Sub

Private Sub btnReadWideReceiverData_Click() 'this button will load receiver data from a text file
'open wide receiver file
Open App.Path & "\wideReceivers.txt" For Input As #3

ctr = 0

Do While Not EOF(3) 'read through entire list until complete
    ctr = ctr + 1
    Input #3, player(ctr), team(ctr), rec(ctr), yards(ctr), tds(ctr) 'sort into 5 parallel arrays
Loop 'loop through file until complete

'alert user that the list was loaded
MsgBox "The list of wide receivers was successfully loaded."
End Sub

Private Sub btnReturnToHome_Click() 'this button will hide current form and return user to start form
    frmStart.Show
    frmWideReciever.Hide
End Sub

Private Sub btnSearchRecWR_Click() 'this button will search receiver data for amount entered and list all players who have caught more than the entered amount
'Dim variables
Dim Found As Boolean
Dim I As Integer
Dim recWanted As String

picResults.Cls 'clear picure box for new results

'assign variables
'user enters amount wanted into an input box
recWanted = InputBox("Enter the number of receptions that you wish to see all of the wide receivers who have caught more than the amount entered.")
Found = False

picResults.Print "The wide receivers who have caught more than "; recWanted; " passes are:"
picResults.Print
picResults.Print "Name"; Tab(25); "Team", "Receptions" 'table headings
picResults.Print "*********************************************************************************************************"

For I = 1 To ctr 'searches the entire list
    If recWanted < rec(I) Then
        Found = True
        picResults.Print player(I); Tab(25); team(I), , rec(I)
    End If
Next I

If (Not Found) Then 'if amount is not found this message will be displayed
    MsgBox "No receiver has caught more than " & recWanted & " passes."
End If
End Sub

Private Sub btnSearchTdsWR_Click() 'this button will find all receivers who have scored more than the entered amount of tds
'Dim variables
Dim Found As Boolean
Dim I As Integer
Dim tdsWanted As String

picResults.Cls 'clear picture box for new results

'assign variables
'user will enter amount of touchdowns to search for in an input box
tdsWanted = InputBox("Enter the number of touchdowns that you wish to see all of the wide recievers who have scored more than the amount entered.")
Found = False

picResults.Print "The wide receivers who have scored more than "; tdsWanted; " touchdowns are:"
picResults.Print
picResults.Print "Name"; Tab(25); "Team", "Touchdowns"
picResults.Print "*****************************************************************************************************"

For I = 1 To ctr 'searches the entire list
    If tdsWanted < tds(I) Then
        Found = True
        picResults.Print player(I); Tab(25); team(I), , tds(I)
    End If
Next I

If (Not Found) Then 'if amount is not found this message will be displayed
    MsgBox "No receiver has scored more than " & tdsWanted & " touchdowns."
End If
End Sub

Private Sub btnSearchYardsWR_Click() 'this button will find all the receivers that have gained more than the entered amount of yards
'Dim variables
Dim Found As Boolean
Dim I As Integer
Dim yardsWanted As String

picResults.Cls 'clear picture box for new results

'assign variables
'user will enter amount of yards to search for in an input box
yardsWanted = InputBox("Enter an amount of yards that you wish to see all of the wide receivers who have gained more than the amount entered.")
Found = False

picResults.Print "The wide receivers who have gained more than "; yardsWanted; " yards are:"
picResults.Print
picResults.Print "Name"; Tab(25); "Team", "Yards"
picResults.Print "*******************************************************************************************"

For I = 1 To ctr 'searches the entire list
    If yardsWanted < yards(I) Then
        Found = True
        picResults.Print player(I); Tab(25); team(I), , yards(I)
    End If
Next I

If (Not Found) Then 'if amount is not found this message will be displayed
    MsgBox "No player has received for more than the amount of yards that you entered."
End If
End Sub

Private Sub btnSortByReceptions_Click() 'this button will list the receivers by amount of receptions in descending order
Dim TempName As String
Dim TempTeam As String
Dim TempRec As Single
Dim Pass As Integer
Dim I As Integer
Dim Pos As Integer

picResults.Cls 'clear picture box for new results

For Pass = 1 To ctr - 1 'begin bubble search to sort players
    For Pos = 1 To ctr - Pass 'keep track of how many comparasions
        If rec(Pos) < rec(Pos + 1) Then
            TempRec = rec(Pos) 'swaps values if out of order
            rec(Pos) = rec(Pos + 1)
            rec(Pos + 1) = TempRec
            
            TempName = player(Pos) 'swaps values if out of order
            player(Pos) = player(Pos + 1)
            player(Pos + 1) = TempName
            
            TempTeam = team(Pos) 'swaps values if out of order
            team(Pos) = team(Pos + 1)
            team(Pos + 1) = TempTeam
        End If
    Next Pos 'move to next position in list
Next Pass 'do another pass

picResults.Print "Wide receivers sorted by number of receptions:"
picResults.Print
picResults.Print "Player"; Tab(25); "Team", "Receptions"
picResults.Print "*********************************************************************"
For I = 1 To ctr
    picResults.Print I; player(I); Tab(25); team(I), , rec(I)
Next I
End Sub

Private Sub btnSortByTds_Click() 'this button will list receievers by number of touchdowns in descending order
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

picResults.Print "Wide receivers sorted by the number of touchdowns scored:"
picResults.Print
picResults.Print "Player"; Tab(25); "Team", "Touchdowns"
picResults.Print "*********************************************************************"
For I = 1 To ctr
    picResults.Print I; player(I); Tab(25); team(I), , tds(I)
Next I
End Sub

Private Sub btnSortByYards_Click() 'this button lists receivers by amount of yards gained in descending order
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

picResults.Print "Wide receivers sorted by amount of yards gained:"
picResults.Print
picResults.Print "Player"; Tab(25); "Team", "Yards"
picResults.Print "*********************************************************************"
For I = 1 To ctr
    picResults.Print I; player(I); Tab(25); team(I), , yards(I)
Next I
End Sub

Private Sub btnViewReceiverStats_Click() 'this button displays receivers by rank from original list
Dim I As Integer

picResults.Cls 'clear picture box for new results

I = 0

picResults.Print "Name"; Tab(25); "Team", "Receptions", "Yards", "Touchdowns" 'headings for table
picResults.Print "***********************************************************************************"
For I = 1 To ctr 'prints original list
    picResults.Print I; player(I); Tab(25); team(I), , rec(I), yards(I), tds(I)
Next I 'goes to next player
End Sub
