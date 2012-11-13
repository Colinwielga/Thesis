VERSION 5.00
Begin VB.Form frmCharacters 
   BackColor       =   &H00004000&
   Caption         =   "Characters!"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   12255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   9120
      TabIndex        =   11
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Ready to Go!"
      Height          =   495
      Left            =   10320
      TabIndex        =   10
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtCharacter 
      Height          =   495
      Left            =   8280
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
   End
   Begin VB.PictureBox picBio 
      Height          =   1215
      Left            =   5160
      ScaleHeight     =   1155
      ScaleWidth      =   6435
      TabIndex        =   7
      Top             =   1440
      Width           =   6495
   End
   Begin VB.PictureBox picshot 
      Height          =   3375
      Left            =   2640
      ScaleHeight     =   3315
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdWizard 
      Caption         =   "Wizard"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdGiant 
      Caption         =   "Giant"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrincess 
      Caption         =   "Princess"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdKnight 
      Caption         =   "Knight"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblCharacter 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      Caption         =   "Enter in the number for which character you would like to be : 1=Knight 2=Princess 3=Giant 4=Wizard"
      Height          =   735
      Left            =   5160
      TabIndex        =   9
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on the Character's Name for Their Picture and Biography"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   6735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Meet Your Chatacters!"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frmCharacters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'This form is the form that you get to pick your own character
'You choose from 4 characters
'There are bios about each character
'you get to pick one based on the bios you likes



Private Sub cmdGiant_Click()
'This provides the bio and picture for the Giant
    picshot.Cls
    picBio.Cls
    picshot.Picture = LoadPicture(App.Path & "\giantshot.jpg")
    Dim k As Integer
    For k = 1 To Ctr
        If Names(k) = "Giant" Then
            picBio.Print "Strengths:"
            picBio.Print Strengths(k)
            picBio.Print "_____________________________________________________________________"
            picBio.Print "Weaknesses:"
            picBio.Print Weaknesses(k)
        End If
    Next k
End Sub

Private Sub cmdKnight_Click()
'Provides the bio and picture for the knight
    picshot.Cls
    picBio.Cls
    picshot.Picture = LoadPicture(App.Path & "\knightshot.jpg")
    Dim f As Integer
    For f = 1 To Ctr
        If Names(f) = "Knight" Then
            picBio.Print "Strengths:"
            picBio.Print Strengths(1)
            picBio.Print "_____________________________________________________________________"
            picBio.Print "Weaknesses:"
            picBio.Print Weaknesses(1)
        End If
    Next f
End Sub

Private Sub cmdPrincess_Click()
 'provides the bio and picture for the princess
    picshot.Cls
    picBio.Cls
     picshot.Picture = LoadPicture(App.Path & "\princessshot.jpg")
    Dim d As Integer
    For d = 1 To Ctr
        If Names(d) = "Princess" Then
         picBio.Print "Strengths:"
         picBio.Print Strengths(d)
         picBio.Print "_____________________________________________________________________"
         picBio.Print "Weaknesses:"
         picBio.Print Weaknesses(d)
        End If
    Next d
End Sub

Private Sub CmdQuit_Click()
    End
End Sub

Private Sub cmdStart_Click()
    'Takes you to the form you want next
    CharacterName = InputBox("Enter in the name of your character:", "Character's Name")
    If txtCharacter.Text = 1 Then
        Character = "Knight"
        frmCharacters.Hide
        frmKnightStart.Show
        MsgBox "Welcome Sir " & CharacterName & "! Ready to Begin?", , "Welcome Again!"
    ElseIf txtCharacter.Text = 2 Then
        Character = "Princess"
        frmCharacters.Hide
        frmprincessstart.Show
        MsgBox "Welcome Princess " & CharacterName & "! Ready to Begin?", , "Welcome Again!"
    ElseIf txtCharacter.Text = 3 Then
        Character = "Giant"
        frmCharacters.Hide
        MsgBox "You died of your own stench.", , "Gross!"
        MsgBox "This is where your story ends. Start over.", , "Story Ends"
        frmCharacters.Hide
        frmWelcome.Show
        picshot.Cls
        picBio.Cls
    ElseIf txtCharacter.Text = 4 Then
        Character = "Wizard"
        frmCharacters.Hide
        frmWizardStart.Show
        MsgBox "Welcome Child " & CharacterName & ", of Mystic Forest! Ready to Begin?", , "Welcome Again!"
    Else
        MsgBox "Ooopsies! Enter in a valid Number Please!", , "Error!"
    End If

End Sub

Private Sub cmdWizard_Click()
'Provides picture and bio for the wizard
picshot.Cls
    picBio.Cls
    picshot.Picture = LoadPicture(App.Path & "\wizardshot.jpg")
    Dim J As Integer
    For J = 1 To Ctr
        If Names(J) = "Wizard" Then
            picBio.Print "Strengths:"
            picBio.Print Strengths(J)
            picBio.Print "_____________________________________________________________________"
            picBio.Print "Weaknesses:"
            picBio.Print Weaknesses(J)
        End If
    Next J
End Sub


