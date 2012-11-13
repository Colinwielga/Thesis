VERSION 5.00
Begin VB.Form frmSuperBowl 
   BackColor       =   &H00800080&
   Caption         =   "The History Of the Super Bowl by Jon Fleischacker"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H00E0E0E0&
      Height          =   7695
      Left            =   5400
      ScaleHeight     =   7635
      ScaleWidth      =   4515
      TabIndex        =   7
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton cmdTeam 
      Caption         =   "Select a Team"
      Height          =   1335
      Left            =   2880
      TabIndex        =   6
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdMToL 
      Caption         =   "Sort By The Number of Wins Per Team"
      Height          =   1335
      Left            =   480
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdWETH 
      Caption         =   "How Many Super Bowl Wins Per Team"
      Height          =   1335
      Left            =   2880
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select a Year"
      Height          =   1335
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1335
      Left            =   3240
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   2
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort By Super Bowl XXVII To Super Bowl I"
      Height          =   1335
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H80000000&
      Caption         =   " Super Bowl History"
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblPic 
      Caption         =   "Trent Dilfer Holding the Super Bowl XXV Trophy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   1725
      Left            =   720
      Picture         =   "NFL.frx":0000
      Top             =   6240
      Width           =   1410
   End
End
Attribute VB_Name = "frmSuperBowl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SuperBowl (jjfleischacker NFL.vbp)
'Form Name: frmSuperBowl (NFL.frm)
'Author: Jon Fleischacker
'Date Written: October 28, 2003
'Purpose: To get Super Bowl information that includes all the
    'scores of each team of each Super Bowl. Also, it prints
    'information of who won the most Super Bowls to who won the
    'least amount of Super Bowls.

'Option Explicit is a command to force
'the user to declare all variables before they can be used.
Option Explicit
'This is an array of an integer that is declared globally
Dim Year(1 To 50) As Integer
'This is an array of strings that are declared globally
Dim strNameW(1 To 50) As String
'This is an array of strings that are declared globally
Dim strNameL(1 To 50) As String
'This is an array of an integer that is declared globally
Dim WScore(1 To 50) As Integer
'This is an array of an integer that is declared globally
Dim LScore(1 To 50) As Integer
'This is an integer that is declared globally
Dim i As Integer
'This is a string that is declared globally
Dim srtPath As String
'This is an array of strings that are declared globally
Dim strNameSuper(1 To 50) As String
'This is an integer that is declared globally
Dim Wins(1 To 50) As Integer



Private Sub cmdRead_Click()
'It clears whatever is in pbxResults
pbxResults.Cls
For i = 1 To 37 'number of passes through the list
    'Prints results that the user entered
    pbxResults.Print Year(i); Tab(10); strNameW(i); Tab(30); WScore(i); Tab(35); strNameL(i); Tab(55); LScore(i)
Next i 'Checks to see if the number is between 1 and 37
Close #1 'Close the file used for input
End Sub

Private Sub cmdSort_Click()
'Declares variable to be read locally
Dim pass As Integer
'Declares variable to be read locally
Dim temp As Integer
'Declares variable to be read locally
Dim temp2 As String
'Declares variable to be read locally
Dim N As Single
N = 37
'It clears whatever is in pbxResults
pbxResults.Cls
For pass = 1 To N - 1 'number of passes through the list
    For i = 1 To N - pass 'Number of comparisons for each pass.
        If Year(i) < Year(i + 1) Then 'compare adjacent names
            temp = Year(i + 1) 'swap if necessary
            Year(i + 1) = Year(i)
            Year(i) = temp
            temp2 = strNameW(i + 1) 'swap if necessary
            strNameW(i + 1) = strNameW(i)
            strNameW(i) = temp2
            temp = WScore(i + 1) 'swap if necessary
            WScore(i + 1) = WScore(i)
            WScore(i) = temp
            temp2 = strNameL(i + 1) 'swap if necessary
            strNameL(i + 1) = strNameL(i)
            strNameL(i) = temp2
            temp = LScore(i + 1) 'swap if necessary
            LScore(i + 1) = LScore(i)
            LScore(i) = temp
        End If 'ends if all the comparisons are made adjacent to one another
    Next i 'checks to see if all comparisons are made
Next pass 'passes throught the list until 37
For i = 1 To N 'number of passes through the list
    'Prints results that the user entered
    pbxResults.Print Year(i); Tab(10); strNameW(i); Tab(30); WScore(i); Tab(35); strNameL(i); Tab(55); LScore(i)
Next i
End Sub

Private Sub cmdQuit_Click()
    'End this program immediately.
    End
End Sub



Private Sub cmdSelect_Click()
'Declares variable to be read locally
Dim iSelect As Integer
'Declares variable to be read locally
Dim Found As Boolean
'Declares variable to be read locally
Dim N As Integer
Found = False
i = 0
N = 37
iSelect = InputBox("Enter a Year")
'It clears whatever is in pbxResults
pbxResults.Cls
Do While i <= N - 1 And Found = False
    i = i + 1
    If iSelect = Year(i) Then
        Found = True 'name = NFL Team
    End If 'End if team is an NFL team
Loop 'checks to see if the number is between 1 and 37
If Found = True Then 'If name is an NFL team
    'Prints results that the user entered
    pbxResults.Print strNameW(i); " "; WScore(i); ","; " "; strNameL(i); " "; LScore(i)
Else 'If not true
    'Prints results that the user entered
    pbxResults.Print "The Super Bowl was not played in the year you selected"
End If
End Sub

Private Sub cmdMToL_Click()
'Declares variable to be read locally
Dim pass As Integer
'Declares variable to be read locally
Dim temp As Integer
'Declares variable to be read locally
Dim temp2 As String
'Declares variable to be read locally
Dim N As Single
N = 34
'It clears whatever is in pbxResults
pbxResults.Cls
For pass = 1 To N - 1 'number of passes through the list
    For i = 1 To N - pass 'Number of comparisons for each pass.
        If Wins(i) < Wins(i + 1) Then 'compare adjacent names
            temp = Wins(i + 1) 'swap if necessary
            Wins(i + 1) = Wins(i)
            Wins(i) = temp
            temp2 = strNameSuper(i + 1) 'swap if necessary
            strNameSuper(i + 1) = strNameSuper(i)
            strNameSuper(i) = temp2
        End If 'ends if all camparisone are made adjacent to one another
    Next i 'checks to see if all comparisons are made
Next pass 'passes through the list until 34
For i = 1 To 34 'number of passes through the list
    'Prints results that the user entered
    pbxResults.Print strNameSuper(i); Tab(25); Wins(i)
Next i 'checks to see if the number is between 1 and 34
End Sub

Private Sub cmdTeam_Click()
'Declares variable to be read locally
Dim iTeam As String
'Declares variable to be read locally
Dim Found As Boolean
'Declares variable to be read locally
Dim N As Integer
Found = False
i = 0
N = 34
'asks the user to enter a team
iTeam = InputBox("Enter a Team")
'It clears whatever is in pbxResults
pbxResults.Cls
Do While i <= N - 1 And Found = False
    i = i + 1
    If iTeam = strNameSuper(i) Then
        Found = True 'name user entered is an NFL team
    End If 'End if team is an NFl team
Loop 'checks to see if the number is between 1 and 34
If Found Then 'If true
    'Prints results that the user entered
    pbxResults.Print strNameSuper(i); " "; "has"; " "; Wins(i); " "; "Super Bowl wins"
Else 'if not true
    'Prints results that the user entered
    pbxResults.Print "The name you have enter is not in the NFL or be more specific"
    pbxResults.Print "(example: do not type in Vikings. Type Minnesota)"
End If
End Sub

Private Sub cmdSortWins_Click()
'Declares variable to be read locally
Dim pass As Integer
'Declares variable to be read locally
Dim temp As Integer
'Declares variable to be read locally
Dim temp2 As String
'Declares variable to be read locally
Dim N As Single
N = 34
'It clears whatever is in pbxResults
pbxResults.Cls
For pass = 1 To N - 1 'number of passes through the list
    For i = 1 To N - pass 'Number of comparisons for each pass.
        If Wins(i) < Wins(i + 1) Then 'compare adjacent names
            temp = Wins(i + 1) 'swap if necessary
            Wins(i + 1) = Wins(i)
            Wins(i) = temp
            temp2 = strNameSuper(i + 1)
            strNameSuper(i + 1) = strNameSuper(i)
            strNameSuper(i) = temp2
        End If 'ends if all swaps are made
    Next i 'checks to see if all comparisons are made
Next pass 'passes through the list until 34
For i = 1 To 34 'number of passes through the list
    'Prints results that the user entered
    pbxResults.Print strNameSuper(i); Tab(25); Wins(i)
Next i 'checks to see if the number is between 1 and 34
End Sub



Private Sub cmdWETH_Click()
'It clears whatever is in pbxResults
pbxResults.Cls
For i = 1 To 34 'number of passes through the list
    'Prints results that the user entered
    pbxResults.Print strNameSuper(i); Tab(25); Wins(i)
Next i 'Chacks to see if the number is between 1 and 34
End Sub



Private Sub Form_Load()
'Prepare the file to be read
Open strPath & "VBProject.txt" For Input As #1
For i = 1 To 37 'number of passes through the list
    'Reads the file
    Input #1, Year(i), strNameW(i), strNameL(i), WScore(i), LScore(i)
Next i 'checks to see if the number si between 1 and 37
Close #1 'Close the file used for input
'Prepare the file to be read
Open strPath & "VBProject2.txt" For Input As #2
For i = 1 To 34 'number of passes through the list
    Input #2, strNameSuper(i), Wins(i) 'Reads the file
Next i 'checks to see if the number is between 1 and 34
Close #2 'Close the file used for input
'Prepares file to be read
strPath = "N:\CS130\jjfleischacker\"
End Sub

