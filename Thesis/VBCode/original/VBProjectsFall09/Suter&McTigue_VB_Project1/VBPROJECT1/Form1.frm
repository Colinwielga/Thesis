VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   Caption         =   "Form1"
   ClientHeight    =   10455
   ClientLeft      =   4680
   ClientTop       =   2415
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   10455
   ScaleWidth      =   14550
   Begin VB.CommandButton cmdPosition 
      Caption         =   "Select Position"
      Height          =   735
      Left            =   7920
      TabIndex        =   10
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdTicket 
      Caption         =   "To Check out Ticket Info Click Here"
      Height          =   735
      Left            =   840
      TabIndex        =   9
      Top             =   7320
      Width           =   3855
   End
   Begin VB.CommandButton cmdStadium 
      Caption         =   "To Check out Stadium Info Click Here"
      Height          =   735
      Left            =   840
      TabIndex        =   8
      Top             =   8280
      Width           =   3855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Picture"
      Height          =   735
      Left            =   10560
      TabIndex        =   7
      Top             =   3960
      Width           =   2415
   End
   Begin VB.PictureBox picNumber 
      Height          =   3735
      Left            =   9480
      ScaleHeight     =   3675
      ScaleWidth      =   5355
      TabIndex        =   6
      Top             =   120
      Width           =   5415
   End
   Begin VB.TextBox txtNumber 
      Height          =   615
      Left            =   8040
      TabIndex        =   4
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "Select Player you want to view"
      Height          =   855
      Left            =   6480
      TabIndex        =   3
      Top             =   7920
      Width           =   2535
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Quit"
      Height          =   1215
      Left            =   12240
      TabIndex        =   2
      Top             =   9000
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   5895
      Left            =   6240
      ScaleHeight     =   5835
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton cmdPos 
      Caption         =   "Find Player"
      Height          =   735
      Left            =   6360
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblNumber 
      Caption         =   "Enter the Number of the player that you would like to view"
      Height          =   615
      Left            =   6240
      TabIndex        =   5
      Top             =   7200
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lastname(1 To 100) As String, firstname(1 To 100) As String, Position(1 To 100) As String

'USAsoccer
'Form1
'author Marty
'October 14
'This is the main form used to load an array of USA soccer names and maneuver through the rest of the project
'it is also used to see player pictures

Private Sub cmdclear_Click()
'this clears the picture from the picture box
picNumber.Picture = LoadPicture("")
picresults.Cls
End Sub

Private Sub cmdEnd_Click()
'this is used to quit
End
End Sub

Private Sub cmdNumber_Click()
Dim pictureNumber As Integer
 'this is a do while fuction that uses the entered textbox number to find corrisponding picture
pictureNumber = txtNumber.Text
Do While (pictureNumber < 1 Or pictureNumber > 58)
    'this is if the user enters invalid number
    pictureNumber = InputBox("Enter an integer in the range 1 to 58")
Loop

picNumber.Picture = LoadPicture(App.Path & "\SoccerPictures\" & names(pictureNumber))


End Sub

Private Sub cmdPos_Click()
'this is an array to find soccer names and positions
Open App.Path & "\SoccerNames.txt" For Input As #1

ctr = 0

Do Until EOF(1)
    ctr = ctr + 1
    Input #1, lastname(ctr), firstname(ctr), Position(ctr)
Loop
'brings user to the position button to find player based on position
MsgBox ("Click on the Select Position Button")

End Sub

Private Sub cmdPosition_Click()
Dim Posit As String, i As Integer
'this is the for next that separates the array by position and loads the info to picture box
Posit = InputBox("Enter the Positions of the Player you would like to view: Midfielder, Forward, Defender, or Goalkeeper .")
For i = 1 To ctr
    If Position(i) = Posit Then
        picresults.Print i; "."; lastname(i); ", "; firstname(i)
    End If
Next i
End Sub

Private Sub cmdStadium_Click()
'switches to stadium form
Form1.Hide
frmsoccerstadiums.Show
End Sub

Private Sub cmdTicket_Click()
'switches to the ticket form
Form1.Hide
frmticketinfo.Show
End Sub



