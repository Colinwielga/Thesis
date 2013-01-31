VERSION 5.00
Begin VB.Form FormToyStory
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   12915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   22395
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   Picture         =   "FormToyStory.frx":0000
   ScaleHeight     =   12915
   ScaleWidth      =   22395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPicGuessStart
      BackColor       =   &H0000FF00&
      Caption         =   "Start/Next"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdHome
      Caption         =   "Back to Games"
      Height          =   1095
      Left            =   15120
      TabIndex        =   18
      Top             =   11640
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit
      Caption         =   "Exit"
      Height          =   1095
      Left            =   18240
      TabIndex        =   17
      Top             =   11640
      Width           =   2535
   End
   Begin VB.PictureBox picWoodyArm
      Height          =   4575
      Left            =   6600
      Picture         =   "FormToyStory.frx":24EA42
      ScaleHeight     =   4515
      ScaleWidth      =   4995
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.PictureBox picWoodyFull
      Height          =   7575
      Left            =   6240
      Picture         =   "FormToyStory.frx":2971EC
      ScaleHeight     =   7515
      ScaleWidth      =   5955
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.PictureBox picRexNeck
      Height          =   4575
      Left            =   6480
      Picture         =   "FormToyStory.frx":3299EE
      ScaleHeight     =   4515
      ScaleWidth      =   5235
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.PictureBox picRexFull
      Height          =   5535
      Left            =   6360
      Picture         =   "FormToyStory.frx":377AC8
      ScaleHeight     =   5475
      ScaleWidth      =   5475
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.PictureBox picHammySlot
      Height          =   6735
      Left            =   6120
      Picture         =   "FormToyStory.frx":3DA84A
      ScaleHeight     =   6675
      ScaleWidth      =   6435
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.PictureBox picHammyFull
      Height          =   6255
      Left            =   6240
      Picture         =   "FormToyStory.frx":4632CC
      ScaleHeight     =   6195
      ScaleWidth      =   7035
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.PictureBox picPotatoHeadEar
      Height          =   6135
      Left            =   6480
      Picture         =   "FormToyStory.frx":4F9566
      ScaleHeight     =   6075
      ScaleWidth      =   5715
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.PictureBox picPotatoHeadFull
      Height          =   8775
      Left            =   5880
      Picture         =   "FormToyStory.frx":5A7270
      ScaleHeight     =   8715
      ScaleWidth      =   6795
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.PictureBox picBuzzFull
      Height          =   6015
      Left            =   6480
      Picture         =   "FormToyStory.frx":66D642
      ScaleHeight     =   5955
      ScaleWidth      =   5355
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.PictureBox picBuzzArm
      Height          =   4335
      Left            =   6240
      Picture         =   "FormToyStory.frx":6D5B44
      ScaleHeight     =   4275
      ScaleWidth      =   5835
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.CommandButton cmdRex
      BackColor       =   &H00FFFF80&
      Caption         =   "Rex"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   13560
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdHammy
      BackColor       =   &H00FFFF80&
      Caption         =   "Hammy"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   13560
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdPotatoHead
      BackColor       =   &H00FFFF80&
      Caption         =   "Mr. Potato Head"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   13560
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdWoody
      BackColor       =   &H00FFFF80&
      Caption         =   "Woody"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13560
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdBuzz
      BackColor       =   &H00FFFF80&
      Caption         =   "Buzz Lightyear"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   13560
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdPicGame
      BackColor       =   &H0080FF80&
      Caption         =   "Picture Guess"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Shape Shape2
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   12975
      Left            =   17400
      Top             =   0
      Width           =   3975
   End
   Begin VB.Shape Shape1
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   0
      Top             =   10080
      Width           =   21015
   End
   Begin VB.Label Label1
      BackColor       =   &H000000C0&
      Caption         =   "Choose A Game to Play"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "FormToyStory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Disney Games
'Form Name: FormToyStory
'Author: Jonathan Legnitto
'2/25/10
'Objective: The objective of this form was to make a game where an image that might not be easy to identify because it is out of focus
            'And to have the user click the button of the character they thought it was.  The program then tells them if they are right
            'Or wrong and either shows them the full picture or disabled the incorrect button until they guess correctly

Option Explicit
Dim I As Integer, CtrTS As Integer



Private Sub cmdBuzz_Click()


  Dim aaaa as Single
    If CtrTS = 0 Then                   'This If/Then statement acknowledges a correct/false answer and disables the wrongly attempted button
        picBuzzArm.Visible = False        'I can use this format for each button, but switching the correct answer
        picBuzzFull.Visible = True
        MsgBox ("Correct, Good Job " & Username & "!")

    ElseIf CtrTS = 1 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdBuzz.Enabled = False
    ElseIf CtrTS = 2 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdBuzz.Enabled = False
    ElseIf CtrTS = 3 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdBuzz.Enabled = False
    ElseIf CtrTS = 4 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdBuzz.Enabled = False
    Else


End If


End Sub

Private Sub cmdHammy_Click()

    If CtrTS = 0 Then                   'This If/Then statement acknowledges a correct/false answer and disables the wrongly attempted button
        MsgBox ("Sorry " & Username & ",try again")
        cmdHammy.Enabled = False
    ElseIf CtrTS = 1 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdHammy.Enabled = False
    ElseIf CtrTS = 2 Then
        picHammySlot.Visible = False        'I can use this format for each button, but switching the correct answer
        picHammyFull.Visible = True
        MsgBox ("Correct, Good Job " & Username & "!")

    ElseIf CtrTS = 3 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdHammy.Enabled = False
    ElseIf CtrTS = 4 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdHammy.Enabled = False
    Else
    Dim bbbb as String

End If

End Sub

Private Sub cmdHome_Click()
FormHome.Show 'Brings the user back to the home screen
End Sub

Private Sub cmdPicGame_Click()
MsgBox (Username & ", You will be shown part of one of the characters from toy story.  Click the correct button to reveal the full character, Click the Start/Next Button to begin and to move to the next picture.")

cmdPicGuessStart.Visible = True

CtrTS = -1 'Starts the counter for the game...I had to make it "-1" instead of Zero because My starting value is zero




End Sub

Private Sub cmdPicGuessStart_Click()
'Reveals the buttons and re-enables them all
cmdBuzz.Visible = True
cmdBuzz.Enabled = True
cmdWoody.Visible = True
cmdWoody.Enabled = True
cmdHammy.Visible = True
cmdHammy.Enabled = True
cmdRex.Visible = True
cmdRex.Enabled = True
cmdPotatoHead.Visible = True
cmdPotatoHead.Enabled = True

'Clears all of the full images from the view of the user
picBuzzFull.Visible = False
picWoodyFull.Visible = False
picHammyFull.Visible = False
picPotatoHeadFull.Visible = False
picRexFull.Visible = False


CtrTS = CtrTS + 1

    If CtrTS = 0 Then
        picBuzzArm.Visible = True 'shows the small section of the image
        picWoodyArm.Visible = False
        picHammySlot.Visible = False
        picRexNeck.Visible = False
        picPotatoHeadEar.Visible = False
    ElseIf CtrTS = 1 Then
        picWoodyArm.Visible = True
        picBuzzArm.Visible = False
        picHammySlot.Visible = False
        picRexNeck.Visible = False
        picPotatoHeadEar.Visible = False
    ElseIf CtrTS = 2 Then
        picHammySlot.Visible = True
        picBuzzArm.Visible = False
        picWoodyArm.Visible = False
        picRexNeck.Visible = False
        picPotatoHeadEar.Visible = False
    ElseIf CtrTS = 3 Then
        picRexNeck.Visible = True
         picBuzzArm.Visible = False
        picWoodyArm.Visible = False
        picHammySlot.Visible = False
        picPotatoHeadEar.Visible = False
    ElseIf CtrTS = 4 Then
        picPotatoHeadEar.Visible = True
        picBuzzArm.Visible = False
        picWoodyArm.Visible = False
        picHammySlot.Visible = False
        picRexNeck.Visible = False
        picBuzzArm.Visible = False
        picWoodyArm.Visible = False
        picHammySlot.Visible = False
        picRexNeck.Visible = False
    ElseIf CtrTS > 3 Then
        CtrTS = -1
Else

End If

End Sub

Private Sub cmdPotatoHead_Click()
    If CtrTS = 0 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdPotatoHead.Enabled = False
    ElseIf CtrTS = 1 Then                   'This  If/Then statement acknowledges a correct/false answer and disables the wrongly attempted button
        MsgBox ("Sorry " & Username & ",try again")
        cmdPotatoHead.Enabled = False
    ElseIf CtrTS = 2 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdPotatoHead.Enabled = False
    ElseIf CtrTS = 3 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdPotatoHead.Enabled = False
    ElseIf CtrTS >= 4 Then
        picPotatoHeadEar.Visible = False        'I can use this format for each button, but switching the correct answer
        picPotatoHeadFull.Visible = True
        MsgBox ("Correct, Good Job " & Username & "!")

    Else


End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRex_Click()

   If CtrTS = 0 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdRex.Enabled = False
    ElseIf CtrTS = 1 Then                   'This  If/Then statement acknowledges a correct/false answer and disables the wrongly attempted button
        MsgBox ("Sorry " & Username & ",try again")
        cmdRex.Enabled = False
    ElseIf CtrTS = 2 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdRex.Enabled = False
    ElseIf CtrTS = 3 Then
    picRexNeck.Visible = False        'I can use this format for each button, but switching the correct answer
        picRexFull.Visible = True
        MsgBox ("Correct, Good Job " & Username & "!")
    ElseIf CtrTS >= 4 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdRex.Enabled = False
        cmdRex.Enabled = False
    Else


End If
End Sub

Private Sub cmdWoody_Click()


   If CtrTS = 0 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdWoody.Enabled = False

    ElseIf CtrTS = 1 Then                   'This  If/Then statement acknowledges a correct/false answer and disables the wrongly attempted button
        picWoodyArm.Visible = False        'I can use this format for each button, but switching the correct answer
        picWoodyFull.Visible = True
        MsgBox ("Correct, Good Job " & Username & "!")


    ElseIf CtrTS = 2 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdWoody.Enabled = False
    ElseIf CtrTS = 3 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdWoody.Enabled = False
        cmdWoody.Enabled = False
    ElseIf CtrTS = 4 Then
        MsgBox ("Sorry " & Username & ",try again")
        cmdWoody.Enabled = False
    Else


End If

End Sub
