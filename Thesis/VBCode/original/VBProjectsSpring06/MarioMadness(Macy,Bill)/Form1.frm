VERSION 5.00
Begin VB.Form frmMariocatcher 
   Caption         =   "Mario Catcher"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   492
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   492
   End
   Begin VB.Timer tmrMario 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   3720
   End
   Begin VB.Timer tmrGlobal 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   3720
   End
   Begin VB.PictureBox PicMario 
      Height          =   612
      Left            =   2520
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "By Bill Macy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6600
      TabIndex        =   1
      Top             =   0
      Width           =   1932
   End
   Begin VB.Image imgMario1 
      Height          =   468
      Left            =   2520
      Top             =   2520
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "frmMariocatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Mario Madness
'Form name: frmMarioCatcher
'Author: Bill Macy
'Date Written: Tuesday March 14th, 2006
'Objective of form:  This form allows the user to play a mario catching game.  After selecting a level of dificulty, the user
                'starts the game and clicks on the moving mario picture.  Every correct click is a point.  The total is tallied
                'and the user ends the game by selecting stop from the options.  They can also select a different level of
                'difficulty or view the about section that tells them I made the game.  The last option is to exit the game.

Option Explicit

'the global timer is the time in between when the user clicks on the picture and when the next picture shows up
'the mario timer is how long the picture of mario stays showing

Public Score As Integer     'declares my variables
Public Amount As Integer
Public MarioTimeout As Integer
Public GlobTimeout As Integer
Private GlobTimeoutCount As Integer
Private MarioTimeoutCount As Integer

Private Sub cmdStart_Click()
    Dim success As Single       'declares the variable
    Score = 0       'sets variables equal to zero
    Amount = 0
    success = (1 / 1)       'helps in calculation of your success
    MarioTimeoutCount = 0       'sets the mario timeout count to zero so the timer can go through the loop and you have time to click the picture
    frmMariocatcher.Caption = "MarioCatcher - Score: " & Score & " Success Rate: " & FormatPercent(success)     'displays this information in the title bar as a caption
    tmrGlobal.Enabled = True        'starts the global timer
End Sub

Private Sub cmdStop_Click()
    Dim success As Single       'declares the success variable
    success = (Score / Amount)      'calculates the success by taking how any you clicked by the amount of mario's there were
    tmrGlobal.Enabled = False       'stops the global timer
    tmrMario.Enabled = False        'stops the mario mover timer
    imgMario1.Visible = False       'hides the moving mario picture
    MsgBox frmMariocatcher.Caption, , "Overall Stats"       'displays the overall stats
    MsgBox "Thanks for playing! Try another level or play again soon!", , "Thanks!"     'displays a message asking the user to play again
    frmOptions.Show     'shows the options menu
    frmMariocatcher.Hide        'hides the mario catcher screen
End Sub

Private Sub Form_Initialize()
    Dim TempGlobTimeout As String       'declares my variables
    Dim TempMarioTimeout As String

    imgMario1 = PicMario        'sets the image box to the picture of mario that moves
    If TempGlobTimeout = "" Or TempMarioTimeout = "" Then   'if nothing is selected in the options menu, the default is set with the following information (easiest level)
        GlobTimeout = 10        'sets the global timer to 10
        MarioTimeout = 100      'sets the mario timer to 100
    Else
        TempGlobTimeout = CInt(GetSetting(App.Title, "GameFlags", "GlobTimeout"))       'gets the setting that the user selected and sets the global timer according to those in the options menu
        TempMarioTimeout = CInt(GetSetting(App.Title, "GameFlags", "MarioTimeout"))     'gets the setting that the user selected and sets the mario timer according to those in the options menu
    End If
End Sub

Private Sub imgMario1_Click()
   Dim success As Single        'declares the variable
   
   Score = Score + 1        'for each click on the picture you add one to the score
   success = (Score / Amount)       'divides the score by the amount of marios there actually were to calculate a total
   
   frmMariocatcher.Caption = "MarioCatcher - Score: " & Score & " Success Rate: " & FormatPercent(success)      'displays the overall stats in a caption on the title bar
   tmrMario.Enabled = False     'stops the moving mario timer
   MarioTimeoutCount = 0        'sets the mario timeoutcount equal to zero so the time starts over for you to have the chance to click on the picture
   imgMario1.Visible = False        'makes the picture hide so it can randomly be placed again
   tmrGlobal.Enabled = True     'starts the global timer
   
End Sub


Private Sub tmrGlobal_Timer()
    If GlobTimeoutCount = GlobTimeout Then      'once the level set number equals that of the incremented one, the program will follow this loop
        GlobTimeoutCount = 0        'it resets the globtimeout count
        
        Randomize       'randomizes where the picture will be
        With imgMario1      'takes the image and follows this loop
            .Top = (((frmMariocatcher.ScaleHeight - imgMario1.Height) - 0) * Rnd + 0)       'moves the picture randomly from the top of the form
            .Left = (((frmMariocatcher.ScaleWidth - imgMario1.Width) - 0) * Rnd + 0)        'moves the picture randomly from the left of the form
            .Visible = True     'sets the pictures visibility to true so the image can be seen
        End With
        Amount = Amount + 1     'increments amount by one since the picture has been moved
        tmrMario.Enabled = True     'starts the mario timer
        tmrGlobal.Enabled = False       'stops the global timer since the picture is shown
    Else
        GlobTimeoutCount = GlobTimeoutCount + 1     'increments the globtimeout count by one so you can continue on until it reaches the selected amount from the options menu
    End If
End Sub

Private Sub tmrMario_Timer()
    If MarioTimeoutCount = MarioTimeout Then        'once the mariotimeoutcount is incremented to equl the level that you selected this loop occurs
        Dim success As Single
        success = (Score / Amount)      'calculates the success by dividing the amount of correct clicks by the amount of marios shown
        frmMariocatcher.Caption = "MarioCatcher - Score: " & Score & " Success Rate: " & FormatPercent(success)     'prints that information to the user in the title bar
        MarioTimeoutCount = 0       'sets your timeout count to zero so you can go to the next picture
        tmrGlobal.Enabled = True        'starts the global timer
        imgMario1.Visible = False       'hides the picture
        tmrMario.Enabled = False        'stops the mario timer since the picture is hidden
    Else
        MarioTimeoutCount = MarioTimeoutCount + 1       'adds one to the initial zero until it reaches the amount specified in the level you selected then it will follow the loop above this one
    End If
End Sub


