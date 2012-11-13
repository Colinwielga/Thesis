VERSION 5.00
Begin VB.Form frmHardCase 
   BackColor       =   &H00000000&
   Caption         =   "Hard Case"
   ClientHeight    =   3090
   ClientLeft      =   2160
   ClientTop       =   1950
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdanswer 
      BackColor       =   &H00FF00FF&
      Caption         =   "Click to Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8640
      Width           =   2655
   End
   Begin VB.PictureBox pichard 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   240
      ScaleHeight     =   6855
      ScaleWidth      =   14655
      TabIndex        =   0
      Top             =   1440
      Width           =   14655
   End
   Begin VB.Label lblwatchout 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmHardCase.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "frmHardCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'This form is the hard case file. It reads it and then makes you answer it. There is no going back or forward until you answered it correctly. Not recommended for the non-profile students.

Private Sub cmdAnswer_Click()
'Declare my variables for this button
    Dim Answer As String
    'Asks for input using an input box that will pop up and ask.
    Answer = InputBox("please type in your COMPLETE answer, remember to watch the spelling", "Answer Please")
    'you have to answer it exactly how it appears here. If you don't you can't get out.
    If Answer = "autoerotic manslaughter" Then
        frmHardCase.Hide
        frmHardCaseSolve.Show
    Else
        MsgBox "Sorry you guessed incorrectly", , "Incorrect" 'Displays this until you guess it right
        
    End If
End Sub

'Another form activating display. As soon as the form loads so does my picture box
Private Sub Form_Activate()
'Delcare my variables so i know what i am printing out
Dim ctr As Integer
    pichard.Cls 'Clears any old junk out of my picture box
For ctr = 85 To 105 'The range i am going to want to print
    pichard.Print CaseFile(ctr) 'Prints it line by line
Next ctr 'Until the end of the ctr rage is reached
End Sub

