VERSION 5.00
Begin VB.Form frmFirefighters
   BackColor       =   &H00808000&
   Caption         =   "Members of St. Johns Fire Department"
   ClientHeight    =   5160
   ClientLeft      =   2820
   ClientTop       =   1425
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   8115
   Visible         =   0   'False
   Begin VB.CommandButton cmdLength
      Caption         =   "Search for Firefighter Last Names by Length "
      Enabled         =   0   'False
      Height          =   615
      Left            =   5760
      TabIndex        =   5
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdAlphabetical
      Caption         =   "List the Fire Department Members in Alphabetical Order by Last Name"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5760
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn
      Caption         =   "Return to Main Page"
      Height          =   735
      Left            =   5760
      TabIndex        =   3
      Top             =   4200
      Width           =   2295
   End
   Begin VB.PictureBox picResults
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4515
      ScaleWidth      =   5115
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
   Begin VB.CommandButton cmdSearch
      Caption         =   "Search for a Member"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5760
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdDisplay
      Caption         =   "Display Fire Department Members"
      Height          =   615
      Left            =   5760
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmFirefighters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IDFirst(1 To 50) As String, Position(1 To 50) As String, IDLast(1 To 50) As String
Dim iiii As Long

'Project Name: Saint John's Fire Department
'Form Name: frmFirefighters (this is our roster)
'Authors: JT Trujillo and Matt Mollet
'Date Written: 2/20/2010
'Objective: This form displays three things for the user.
            'It displays our roster by clicking the first button, which
            'reads the file and this displays it in order by rank.
            'The second button allows the user to search the roster
            'for a person by last name.  The third button displays
            'the roster in alphabetic order by last name from A to Z.


Private Sub cmdAlphabetical_Click()

'Declare Variables
Dim pass As Long, pos As Long, j As Long
Dim eeee As String, ffff As String, zzzz As String

picResults.Cls

picResults.Print "Name"; Tab(30); "Position"
picResults.Print "________________________________________"
picResults.Print ""

'sort the names in alphabetic order of last name from A to Z
For pass = 1 To iiii - 1
    For pos = 1 To iiii - pass
        If IDLast(pos) > IDLast(pos + 1) Then
            ffff = IDLast(pos)
            IDLast(pos) = IDLast(pos + 1)
            IDLast(pos + 1) = ffff
            eeee = IDFirst(pos)
            IDFirst(pos) = IDFirst(pos + 1)
            IDFirst(pos + 1) = eeee
            zzzz = Position(pos)
            Position(pos) = Position(pos + 1)
            Position(pos + 1) = zzzz
        End If
    Next pos
Next pass

'Print Results
 For j = 1 To iiii
             picResults.Print IDFirst(j); " "; IDLast(j); Tab(30); Position(j)

    Next j




End Sub

Private Sub cmdDisplay_Click()

'Open file
Open App.Path & "\members.txt" For Input As #1

picResults.Cls

'Display header
picResults.Print "Name"; Tab(30); "Position"
picResults.Print "_______________________________________"
picResults.Print ""

'set iiii to 0
iiii = 0

Do While Not EOF(1)
    iiii = iiii + 1
    'read file
    Input #1, Position(iiii), IDFirst(iiii), IDLast(iiii)
    'Display roster
    picResults.Print IDFirst(iiii); " "; IDLast(iiii); Tab(30); Position(iiii)
Loop

'close file
Close #1

'enable the other two buttons
cmdSearch.Enabled = True
cmdAlphabetical.Enabled = True
cmdLength.Enabled = True
End Sub

Private Sub cmdLength_Click()
Dim SearchLength As Long, K As Long, Found2 As Boolean

'Search for firefighter last names by Length

picResults.Cls
SearchLength = InputBox("To search for firefighters by last name length, please enter a positive number", "Search")
picResults.Print "Last Name", "First Name", "Position "

For K = 1 To iiii
    If SearchLength = Len(IDLast(K)) Then
        Found = True
        picResults.Print IDLast(K), IDFirst(K), Position(K)
    End If
Next K

If Found = False Then
    MsgBox ("Sorry no results found")
End If

End Sub

Private Sub cmdReturn_Click()

'go to main form
frmMain.Visible = True
frmFirefighters.Visible = False


End Sub

Private Sub cmdSearch_Click()
Dim Fireman As String, Found As Boolean, llll As Long, eye As Long

'clear picture box
picResults.Cls
Found = False

'get last name to search for from the user
Fireman = InputBox("Enter the last name of a firefighter.")

picResults.Print "Name"; Tab(30); "Position"
picResults.Print "_____________________________________"
picResults.Print ""

'Search for somebody on the roster who's last name matches the one
'provided by the user
Do While Not Found And eye < iiii
    eye = eye + 1
    For llll = 1 To iiii
    If Fireman = IDLast(llll) Then
    Found = True
    picResults.Print IDFirst(llll); " "; IDLast(llll); Tab(30); Position(llll)
    End If
Next llll
Loop

'Let the user know that there is nobody on the roster that matches
'the name they searched for
If Found = False Then
    MsgBox ("I'm sorry, there is nobody on our roster with a last name of " & Fireman & ".")
End If


End Sub

