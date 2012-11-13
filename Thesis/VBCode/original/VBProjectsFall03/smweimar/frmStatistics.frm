VERSION 5.00
Begin VB.Form frmStatistics 
   BackColor       =   &H00008000&
   Caption         =   "Players Statistics"
   ClientHeight    =   7380
   ClientLeft      =   855
   ClientTop       =   1110
   ClientWidth     =   11520
   ScaleHeight     =   7380
   ScaleWidth      =   11520
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00008000&
      Caption         =   "Refresh Player Information"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   1935
   End
   Begin VB.PictureBox pbxoutput1 
      BackColor       =   &H0000FFFF&
      Height          =   6975
      Left            =   4200
      ScaleHeight     =   6915
      ScaleWidth      =   6795
      TabIndex        =   2
      Top             =   120
      Width           =   6855
   End
   Begin VB.CommandButton cmdpoints 
      BackColor       =   &H00008000&
      Caption         =   "Display players by goals scored (lowest to highest) scored this season"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   3495
   End
   Begin VB.CommandButton cmdHometown 
      BackColor       =   &H00008000&
      Caption         =   "Display Player's Hometown and Picture, by Jersey Number"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   3495
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : MinnesotaWildTeamInformationProgram (Stephanie Weimar's VB Project.vbp)
'Form Name : frmStatistics (frmStatistics.frm)
'Author: Stephanie Weimar
'Date Written: October 29, 2003
'Purpose of Form: To reopen the file that was being used in the previous form.
                 'Then let the user enter a potential jersey number.  If the
                 'number is the number of a player, it will display the number,
                 'the players name and their hometown.  If it is not a number of an actual
                 'player, then a message box will pop up telling the user to try again.
                 'It also give the option of displaying the players in ascending order,
                 'by their number of points scored so far this season.  It also give the user
                 'the option of ending the program.

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
Dim found As Boolean

Private Sub cmdHometown_Click()
    found = False                                           ' To show that if it is found, it does not need to print the phrase
         m = InputBox("Please enter the Number of A Jersey") ' To get the input of a jersey number from the user
                For i = 1 To 24                             ' To search through the list for the number
                    If m = Number(i) Then
                      found = True                          ' When number is found, then clear picture box and display information
                        pbxoutput1.Cls
                        pbxoutput1.Print "Player #"; Number(i); "is "; names(i); " and he is from "; hometown(i); "."
                    End If
                Next i
                    If found = False Then                   ' If the number is not found
                        pbxoutput1.Cls                      ' Clear the picture box
                            MsgBox "Sorry, but this is not the number of a player on the team!  Please try again. ", , "Error" ' Have a messagebox pop up and ask the user to try again
                    End If
End Sub

Private Sub cmdpoints_Click()
    pbxoutput1.Cls                                          ' Clear the picturebox
        For pass = 1 To 23                                  ' To go through the points, and them in descending order along with the players name
            For i = 1 To 24 - pass
                If points(i) > points(i + 1) Then
                    tempnumber = Number(i)
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
        
    pbxoutput1.Print "  Points Scored", "Player Name"       ' To print out the results in the picturebox
    pbxoutput1.Print "So Far This Season"
    pbxoutput1.Print "--------------------------------------------------------------------------------"
        For i = 1 To 24
            pbxoutput1.Print Tab(7); points(i), Tab(29); names(i)
        Next i
            pbxoutput1.Print " "
            pbxoutput1.Print " "
            pbxoutput1.Print " "
            pbxoutput1.Print "With the exception of Manny Fernandez and Dwayne Roloson"
            pbxoutput1.Print "who scored no goals because the play the position of goalie!"
    Close #2                                                ' To close Input #2
End Sub

Private Sub cmdQuit_Click()
    End                                                     ' To end the program
End Sub

Private Sub cmdRefresh_Click()
    Open "wildteam.txt" For Input As #2    ' To open file as Input #2
        For i = 1 To 24                                                 ' Open for all names
            Input #2, names(i), Number(i), hometown(i), points(i)
        Next i
End Sub

Private Sub Form_Load()
    path = "N:\CS130\handin\smweimar\"
End Sub
