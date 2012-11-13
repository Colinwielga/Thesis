VERSION 5.00
Begin VB.Form frmInformation 
   Caption         =   "Information"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5940
   LinkTopic       =   "Form4"
   ScaleHeight     =   5220
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option5 
      Caption         =   "Project Quiz"
      Height          =   615
      Left            =   3240
      TabIndex        =   12
      Top             =   3840
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Pokemon Sort"
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   4080
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Whack a Goblin"
      Height          =   315
      Left            =   4560
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Jasmine"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   4560
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Info."
      Height          =   315
      Left            =   4800
      TabIndex        =   8
      Top             =   4560
      Width           =   855
   End
   Begin VB.PictureBox pbxInfo 
      Height          =   1815
      Left            =   1680
      ScaleHeight     =   1755
      ScaleWidth      =   4035
      TabIndex        =   7
      Top             =   1920
      Width           =   4095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Menu"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   1680
      Picture         =   "Information.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   2880
      Picture         =   "Information.frx":090F
      ScaleHeight     =   915
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.PictureBox Picture7 
      Height          =   1455
      Left            =   4200
      Picture         =   "Information.frx":11BC
      ScaleHeight     =   1395
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox Picture5 
      Height          =   1455
      Left            =   120
      Picture         =   "Information.frx":253A
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   120
      Picture         =   "Information.frx":2FB7
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Designed by Chris Davin"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Click the buttons for info. on the sections."
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   4560
      Width           =   3015
   End
End
Attribute VB_Name = "frmInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : MemoryGamesEtc (Chris Davin's VB Project.vbp)
'Form Name : frmInformation (Information.frm)
'Author: Chris Davin
'Date Written: October 29, 2003
'Purpose of Form: To explain some detail on what each aspect
                 'of the form does.  As well as to provide
                 'some other fun things to click.

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Dim N As String
'This will end the program.
Private Sub cmdQuit_Click()
    End
End Sub
'THis will return you to the main menu.
Private Sub cmdReturn_Click()
    frmInformation.Hide
    frmMainMenu.Show
End Sub

Private Sub Form_Load()

End Sub

'This will reveal info. about the info. section itself.
Private Sub Option1_Click()
    pbxInfo.Cls
    pbxInfo.Print "This is the information section."
    pbxInfo.Print "It provides lots of information on all the sections, beyond"
    pbxInfo.Print "that it's just a fun place to hang out."
End Sub
'This will reveal info. about the Jasmine section.
Private Sub Option2_Click()
    pbxInfo.Cls
    pbxInfo.Print "The Jasmine section is an interactive fortune telling"
    pbxInfo.Print "conversation.  Jasmine is clever,"
    pbxInfo.Print "so you don't want to mess with her."
End Sub
'This will reveal info. about the WhackaGoblin section.
Private Sub Option3_Click()
    pbxInfo.Cls
    pbxInfo.Print "The Whack a Goblin section is a game"
    pbxInfo.Print "much like the popular Whack a Mole game."
    pbxInfo.Print "To win try and clear the screen in as few Whacks"
    pbxInfo.Print "as possible.  Or you can just keep hitting"
    pbxInfo.Print "to your hearts content."
End Sub
'This will reveal info. about the PokemonSort area.
Private Sub Option4_Click()
    pbxInfo.Cls
    pbxInfo.Print "Pokemon are collectable monsters that you can train."
    pbxInfo.Print "Their data is stored on a Pokedex."
    pbxInfo.Print "The Pokemon Sort program helps to sort the information."
End Sub
'This will reveal info. about the Project Quiz section.
Private Sub Option5_Click()
    pbxInfo.Cls
    pbxInfo.Print "The Project Quiz is the final step in this program."
    pbxInfo.Print "The Quiz is given by the mysterious Quiz Master"
    pbxInfo.Print "and all your knowledge is tested."
    pbxInfo.Print "There is no return from the quiz."
End Sub
'Some info. on the picture clicked.
Private Sub Picture1_Click()
    MsgBox "This is Ein, he's really smart.", , "Info."
End Sub
'Info. on the picture clicked.
Private Sub Picture2_Click()
    MsgBox "I have snakes for arms.", , "Info."
End Sub
'Info. on the picture clicked.
Private Sub Picture3_Click()
    MsgBox "I am a carp", , "Info."
End Sub
'Info. on the picture clicked.
Private Sub Picture5_Click()
    MsgBox "Hamtaro is a Hamster", , "Info."
End Sub
'A question from everybodies favorite bear.
Private Sub Picture7_Click()
    N = InputBox("My name is...?", "GRRR")
    If N = "Yogi" Or N = "Yogi Bear" Then
            MsgBox "Yep, youre smarter than the average bear.", , "Reply"
        ElseIf N = "Booboo" Then
            MsgBox "Joke, Joke, Jooooke", , "Reply"
        Else
            MsgBox "Awww, that's not it.", , "Reply"
    End If
End Sub
