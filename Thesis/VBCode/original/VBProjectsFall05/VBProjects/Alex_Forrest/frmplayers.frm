VERSION 5.00
Begin VB.Form frmplayers 
   BackColor       =   &H000040C0&
   Caption         =   "SJU Rugby Players"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit Program"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4440
      TabIndex        =   6
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdgotomain 
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8160
      TabIndex        =   5
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search Players by Position"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7200
      TabIndex        =   4
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdsort2 
      Caption         =   "Sort Players Alphabetically"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7200
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton cmdsort 
      Caption         =   "Sort Positions Alphabetically"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4440
      TabIndex        =   2
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmddisplayplrs 
      Caption         =   "Display Player List"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4440
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.PictureBox picbox 
      Height          =   7695
      Left            =   720
      ScaleHeight     =   7635
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   2535
      Left            =   7200
      Picture         =   "frmplayers.frx":0000
      Top             =   2040
      Width           =   2130
   End
   Begin VB.Image Image1 
      Height          =   2325
      Left            =   4680
      Picture         =   "frmplayers.frx":11ACE
      Top             =   2160
      Width           =   1755
   End
End
Attribute VB_Name = "frmplayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : RugbyVBProject (Rugby.vbp)
'Form Name : frmplayers(frmplayers.frm)
'Author: Alex Forrest
'purpose of the form: This form is designed to go in depth with the players of the
    'SJU rugby team.  One button diplays the roster; another button sorts the players
    'names alphabetically; another button sorts the positions into alphabetical order; and
    'one last button allows the user to input a position and then will display each player
    'on the team that plays that position.

Option Explicit
Dim playerarray(1 To 36) As String
Dim positionarray(1 To 36) As String
Private Sub cmddisplayplrs_Click()
    Dim I As Single
    Dim j As Single
    picbox.Cls 'clears the picture box in case something is already displayed
    
    I = 0 'sets the counter equal to 0
    Open App.Path & "\sjurugbyplayers.txt" For Input As #1 'opens the specified file in path = M:\CSI130\VB Project
    Do Until EOF(1) 'instructs the program to do the following procedure until the end of the file
        I = I + 1 'increments the counter by 1
        Input #1, playerarray(I), positionarray(I) 'sets the whole file equal to input #1
    Loop 'loops to perform the procedure again until the end of the file
    picbox.Print "2005 SJU Fall Rugby Players:"
    picbox.Print
    For j = 1 To I 'begins the loop
        picbox.Print playerarray(j), , positionarray(j) 'prints corresponding parts of each array
    Next j 'goes back to the beginning of loop to print the rest of the array
    Close #1 'closes the file
End Sub

Private Sub cmdgotomain_Click()
    frmplayers.Hide
    frmMainmenu.Show 'takes the user back to the main menu
End Sub

Private Sub cmdquit_Click()
    End 'ends the program
End Sub

Private Sub cmdsearch_Click()
    Dim notfound As Boolean
    Dim I As Integer
    Dim p As String
    p = InputBox("Enter the position you wish to find a list of players for. Enter the position exactly as you see it on the list.(They are case sensitive!)", "Positions") 'sets p equal to whatever is in the input box
    I = 0 'sets I equal to 0
    notfound = True 'sets notfound equal to true
    picbox.Cls 'prints blank line
    picbox.Print "The players that play that position are:" 'prints this line
    picbox.Print
    For I = 1 To 36 'begins the loop to search
        If p = positionarray(I) Then 'searches file to find match
            picbox.Print playerarray(I), positionarray(I) 'if match found, it prints this line
            notfound = False
        End If 'ends if statement
    Next I 'loops to search the next part of the array
    If notfound Then 'if no match is found, it prints the following line
        picbox.Print "Sorry, you did not input a correct position!"
    End If 'ends the if statement
End Sub

Private Sub cmdsort_Click()
    Dim pass As Integer
    Dim X As Integer, r As Integer
    Dim temp As String
    Dim tempop As String
    Dim I As Single
    picbox.Cls
    I = 36 'sets the array equal to 36
    For pass = 1 To I - 1 'begins the sorting process by making the first pass
        For X = 1 To I - pass 'compares the corresponding parts to the array
            If positionarray(X) > positionarray(X + 1) Then 'begins the if statement to compare each part of the array
                temp = playerarray(X) 'uses a holding spot to match up the the correct information
                playerarray(X) = playerarray(X + 1) 'sets each spot to the correct information after being compared and sorted
                playerarray(X + 1) = temp
                tempop = positionarray(X) 'uses a holding spot for the other array to match up the correct info
                positionarray(X) = positionarray(X + 1) 'sets each spot to the correct information after being compared and sorted
                positionarray(X + 1) = tempop
            End If 'ends the if case
        Next X 'peforms the above procedure for the next part of the array
    Next pass 'goes to make the next pass
    picbox.Print "The players in alphabetical position order are:"
    picbox.Print
    For r = 1 To 36 'begins the loop to print both arrays
        picbox.Print playerarray(r); Tab(18); positionarray(r) 'prints both arrays in alphabetical position order
    Next r 'loops to print the next parts
End Sub

Private Sub cmdsort2_Click()
    Dim pass As Integer
    Dim X As Integer, r As Integer
    Dim temp As String
    Dim tempop As String
    Dim I As Single
    picbox.Cls
    I = 36 'sets the array equal to 36
    For pass = 1 To I - 1 'begins the sorting process by making the first pass
        For X = 1 To I - pass 'compares the corresponding parts to the array
            If playerarray(X) > playerarray(X + 1) Then 'begins the if statement to compare each part of the array
                temp = positionarray(X) 'uses a holding spot to match up the the correct information
                positionarray(X) = positionarray(X + 1) 'sets each spot to the correct information after being compared and sorted
                positionarray(X + 1) = temp
                tempop = playerarray(X) 'uses a holding spot to match up the the correct information
                playerarray(X) = playerarray(X + 1) 'sets each spot to the correct information after being compared and sorted
                playerarray(X + 1) = tempop
            End If 'ends the if case
        Next X 'peforms the above procedure for the next part of the array
    Next pass 'goes to make the next pass
    picbox.Print "The players in alphabetical order are:"
    picbox.Print
    For r = 1 To 36 'begins the loop to print both arrays
        picbox.Print playerarray(r); Tab(18); positionarray(r) 'prints both arrays in alphabetical first name order
    Next r 'loops to print the next parts
End Sub
