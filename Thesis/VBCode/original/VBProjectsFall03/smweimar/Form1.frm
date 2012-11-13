VERSION 5.00
Begin VB.Form frmRoster 
   BackColor       =   &H00008000&
   Caption         =   "frmRoster"
   ClientHeight    =   7425
   ClientLeft      =   855
   ClientTop       =   1110
   ClientWidth     =   11640
   ScaleHeight     =   7425
   ScaleWidth      =   11640
   Begin VB.CommandButton cmdproceed 
      BackColor       =   &H00008000&
      Caption         =   "Proceed To More Statistics"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton cmdNumbers 
      BackColor       =   &H00008000&
      Caption         =   "Display the Active Roster - Sorted By Their Jersey Number"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdRoster 
      BackColor       =   &H00008000&
      Caption         =   "Display an Active Roster of Players"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.PictureBox pbxoutput 
      BackColor       =   &H0000FFFF&
      Height          =   6015
      Left            =   5040
      ScaleHeight     =   5955
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   360
      Width           =   6255
   End
End
Attribute VB_Name = "frmRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : MinnesotaWildTeamInformationProgram (Stephanie Weimar's VB Project.vbp)
'Form Name : frmRoster (frmRoster.frm)
'Author: Stephanie Weimar
'Date Written: October 29, 2003
'Purpose of Form: To open the file that contains the information and print out
                 ' the current roster of players.  Then to display the roster of players
                 'starting in ascending order.  There is also a button which allows you to move
                 ' to the next form, which contains more statistics, and a button which allows you
                 'to end the program.

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Public path As String
Dim names(1 To 24) As String
Dim Number(1 To 24) As Integer
Dim hometown(1 To 24) As String
Dim i As Integer
Dim n As Integer
Dim pass As Integer
Dim temp As Integer
Dim m As Integer
Dim tempnumber As Integer
Dim tempnames As String
Dim temphometowns As String
Dim points(1 To 24) As Integer
Dim temppoints As Integer
Private Sub cmdNumbers_Click()
    pbxoutput.Cls                                            ' To clear the outputbox
        For pass = 1 To 23                                   ' To sort the players by their number
          For i = 1 To 24 - pass
              If Number(i) > Number(i + 1) Then
                  tempnumber = Number(i)                     ' To keep the name, hometown, and points along with the correct player
                  tempnames = names(i)
                  temphometowns = hometown(i)
                  temppoints = points(i)
                  Number(i) = Number(i + 1)
                  names(i) = names(i + 1)
                  hometown(i) = hometown(i + 1)
                  points(i) = points(i + 1)
                  hometown(i + 1) = temphometowns
                  Number(i + 1) = tempnumber
                  names(i + 1) = tempnames
                  points(i + 1) = temppoints
              End If
          Next i
        Next pass
            pbxoutput.Print "  Team Number", "Player Name"  ' To print out the results of the number sorting
            pbxoutput.Print "--------------------------------------------------------------------------------"
        For i = 1 To 24
            pbxoutput.Print Tab(7); Number(i), Tab(29); names(i)
            Next i
End Sub

Private Sub cmdproceed_Click()
    frmRoster.Visible = False                       ' To make the third form visible, and hide the other two
    frmMinnesotaWildHockeyTeam.Visible = False
    frmStatistics.Visible = True
    Close #1
End Sub

Private Sub cmdQuit_Click()
                                           ' To end the program
    End
End Sub

Private Sub cmdRoster_Click()
      Open path & "wildteam.txt" For Input As #1  ' To open a file for input
        pbxoutput.Print Tab(10); "Active team members for the Minnesota Wild"               ' To print out the Active team members
        pbxoutput.Print Tab(6); "-------------------------------------------------------------------------------------"
            For i = 1 To 24
                Input #1, names(i), Number(i), hometown(i), points(i)
                     pbxoutput.Print Tab(23); names(i)
            Next i
End Sub

Private Sub Form_Load()
    path = "N:\CS130\handin\smweimar\"
End Sub
