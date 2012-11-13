VERSION 5.00
Begin VB.Form frmNumber 
   Caption         =   "Welcome to the Wild Organization"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   Picture         =   "frmNumber.frx":0000
   ScaleHeight     =   8145
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   360
      TabIndex        =   12
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   6000
      TabIndex        =   11
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdGame 
      Caption         =   "Proceed to Game"
      Height          =   615
      Left            =   3240
      TabIndex        =   10
      Top             =   6960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0FFFF&
      Height          =   1095
      Left            =   720
      ScaleHeight     =   1035
      ScaleWidth      =   6915
      TabIndex        =   9
      Top             =   5400
      Width           =   6975
   End
   Begin VB.CommandButton cmdAvailable 
      Caption         =   "Check For Availability"
      Height          =   615
      Left            =   3240
      TabIndex        =   8
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox txtNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   3840
      Width           =   3495
   End
   Begin VB.PictureBox picKevin 
      Height          =   1815
      Left            =   360
      Picture         =   "frmNumber.frx":A833
      ScaleHeight     =   1755
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox picJacques 
      Height          =   1815
      Left            =   6360
      Picture         =   "frmNumber.frx":D255
      ScaleHeight     =   1755
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox picDoug 
      Height          =   1815
      Left            =   3360
      Picture         =   "frmNumber.frx":FB82
      ScaleHeight     =   1755
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.OLE oleAnthem 
      BackColor       =   &H00C0FFFF&
      Class           =   "Package"
      DisplayType     =   1  'Icon
      Height          =   975
      Left            =   6360
      OleObjectBlob   =   "frmNumber.frx":12686
      SourceDoc       =   "M:\CS130\ChrisAdamsVBProj\music\02 The State of Hockey.mp3"
      TabIndex        =   13
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmNumber.frx":2E8C9E
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   3000
      Width           =   7335
   End
   Begin VB.Label lblJaques 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Jacques Lemaire Head Coach Minnesota Wild"
      Height          =   615
      Left            =   6480
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblDoug 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Doug Risebrough General Manager Minnesota Wild"
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblKev 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Kevin Constantine Head Coach Houston Aeros"
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "frmNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Quest for The Cup~Minnesota Wild Trivia Game

'Author: Chris Adams

'Date: November 2007

'This form is where the user selects their number

Private Sub cmdAvailable_Click()    'This button compares the user input number with numbers currently worn by members of the wild
                                    'If the number is not already being worn then a congratualtions message is shown
                                    'If the number is being warn the player is prompted to select a new number

Jersey = txtNumber.Text             'Sets Jersey equal to user input
    If Jersey >= 0 And Jersey < 100 Then
        Select Case Jersey
            Case 0, 2, 6, 13, 17, 18, 20, 23, 30, 35, 36, 39, 40, 93, 94, 95, 97, 98
                picResults.Print "Congratulations! You will wear the number "; Jersey
                picResults.Print "Good Luck!"
                cmdGame.Visible = True
            Case 44 To 54
                picResults.Print "Congratulations! You will wear the number "; Jersey
                picResults.Print "Good Luck!"
                cmdGame.Visible = True
            Case 56 To 91
                picResults.Print "Congratulations! You will wear the number "; Jersey
                picResults.Print "Good Luck!"
                cmdGame.Visible = True
            Case Else
                picResults.Print "I'm Sorry, but "; Jersey; " is not available at this time."
                picResults.Print "Please hit clear and select another number"
        End Select
    Else
        picResults.Print "I'm sorry, but you have entered an invalid jersey number."
        picResults.Print "Please hit clear and select another number."
    End If
        picResults.Print " "
        picResults.Print "Double Click the Wild Icon to hear the Wild Anthem."
End Sub

Private Sub cmdClear_Click()

    picResults.Cls      'Clears picture box

End Sub


Private Sub cmdGame_Click()

    'This message box displays all the info given by the user and then takes them to the game
    MsgBox ("Congratulations " & PlayerFirst & " " & PlayerLast & ". You will play " & Pos & " and start with our AHL affiliate the Houston Aeros. You will wear the number " & Jersey)
    frmNumber.Hide
    frmGame.Show

End Sub

Private Sub cmdQuit_Click()
frmNumber.Hide
frmSources.Show
End Sub
