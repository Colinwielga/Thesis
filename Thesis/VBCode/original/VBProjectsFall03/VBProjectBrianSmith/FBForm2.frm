VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00008000&
   Caption         =   "Form2"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   LinkTopic       =   "Form2"
   ScaleHeight     =   4935
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPoints 
      Caption         =   "Calculate how many points they scored this week"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdDisplayInfo 
      Caption         =   "Click to see player information"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   4320
      Width           =   735
   End
   Begin VB.PictureBox picResults 
      Height          =   3855
      Left            =   2640
      ScaleHeight     =   3795
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   360
      Width           =   4095
   End
   Begin VB.TextBox txtReceiverName 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label LblPLayer 
      Caption         =   "Please enter the receivers name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Yards As Integer, Y As Integer
Public Path As String
    'Declare variables
    'Brian Smith
    'Project1(Football Project)
    'Form1(FB Form1)
    'Oct 26th, 2003
    'Purpose: To calculate the number of points a receiver has acumulated for fantasy football
    


Private Sub cmdStart_Click()
    Form1.Visible = True
    Form2.Visible = False
        'Switches from form 1 to form 2
        'button on form1
End Sub

Private Sub cmdBack_Click()
    Form1.Show
    Form2.Hide
        'Switches from form 2 to form 1
End Sub

Private Sub cmdDisplayInfo_Click()
picResults.Cls
    'Clears the print window
Dim C As Integer
Dim Name As String
Dim Avg As Single
Dim Catches As Integer
    'Declare variables

Name = txtReceiverName.Text
    'Input data from textbox
If Name = "" Then
    MsgBox "Must enter name of player" 'Promtps a box requiring a name
End If
If Name <> "" Then 'Will not allow the program to execute unless a name is enter in the text box
    
C = InputBox("Enter the number of catches", "Catches")
Catches = C
    'Prompts an inputbox to the User, if the value inputed is "-", then a msgbox appears
    
Y = InputBox("Enter the yards gained from the game", "Yards")
Yards = Y
    'Prompts an inputbox to the User, if the value inputed is "-", then a msgbox appears
    
If Y > 0 And C > 0 Then
    Avg = Y / C
    Else: End
End If
    'Finds the yards per catch, if either Y or C are 0, then the program quits
    
picResults.Print Name
picResults.Print "Had"; C; "Catches for"; Y; "yards in the game today"
picResults.Print "Averaging "; FormatNumber(Avg, 2); " yards per catch"
    'Prints Bio of the player from the game
    
cmdPoints.Enabled = True
End If
End Sub

Private Sub cmdPoints_Click()
Dim TD As Integer
Dim Points(1 To 6) As Integer
Dim Range(1 To 6) As Integer
Dim Temp As Integer
Dim I As Integer
Dim Total As Integer
Dim Score As Integer
    'Declares Variables
    
TD = InputBox("Did They score a TD?  If so, how many?")
TD = TD * 6
    'determines the points for TD's

Open Path & "PointScale.txt" For Input As #1
    'Open's up the scale file to File an Array

For I = 1 To 6
    Input #1, Range(I), Points(I)
Next I
    'Fills values from Outside source into an Array
    
For I = 1 To 6
    Temp = Range(I)
    If Temp <= Y Then
        Score = Points(I)
        Total = TD + Score
        Else
        Total = TD + Score
    End If
Next I
    'Compares the value of Y to the the Ranges in the Array, then gives a point value.
Close

picResults.Print "With"; Score; "points for yards and"; TD; "points for Touchdowns,"
picResults.Print "His total points for the week are"; Total
    'Prints the number of points the player has accumulated

cmdPoints.Enabled = False

    
End Sub

Private Sub cmdQuit_Click()
End
   'Quits the VB program
End Sub

Private Sub Form_Load()
cmdPoints.Enabled = False
Path = "N:\CS130\handin\VBProjectBrianSmith\"

End Sub
