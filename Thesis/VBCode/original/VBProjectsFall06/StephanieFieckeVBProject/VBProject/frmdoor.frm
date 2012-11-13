VERSION 5.00
Begin VB.Form frmdoor 
   BackColor       =   &H000040C0&
   Caption         =   "Pick A Door"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   Picture         =   "frmdoor.frx":0000
   ScaleHeight     =   5445
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmain 
      BackColor       =   &H008080FF&
      Caption         =   "Main Menu"
      Height          =   735
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00004080&
      Caption         =   "How About This Door?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmddoor2 
      BackColor       =   &H00004080&
      Caption         =   "Or This Door?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmddoor1 
      BackColor       =   &H00004080&
      Caption         =   "Pick This Door?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmdoor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'A Day For Fun
    'Door
    'Stephanie Fiecke
    '10-31=06
    'This form is entertainment based for the user to simply click on a door and win or not

Option Explicit
Dim winningdoor As Integer


Private Sub cmddoor1_Click()

    'generates a random number and resets itself each time the user clicks on the door
winningdoor = CInt(Int((3 * Rnd()) + 1))

    'If the door is equal to one then the user wins, otherwise it will ask the user to try again
    If winningdoor = 1 Then
        MsgBox "You Win!", vbExclamation, "Congratulations!"
    Else
        MsgBox "Sorry, Wrong Door!", vbCritical, "Try Again!"
    End If
       
     
End Sub

Private Sub cmddoor2_Click()

    'generates random number
winningdoor = CInt(Int((3 * Rnd()) + 1))

    If winningdoor = 1 Then
        MsgBox "You Win!", vbExclamation, "Congratulations!"
    Else
        MsgBox "Sorry, Wrong Door!", vbCritical, "Try Again!"
    End If

End Sub

Private Sub cmdmain_Click()
    'hides the door form and shows the main form
frmdoor.Hide
frmmain.Show
End Sub

Private Sub Command3_Click()

    'generates random number
winningdoor = CInt(Int((3 * Rnd()) + 1))
    If winningdoor = 1 Then
        MsgBox "You Win!", vbExclamation, "Congratulations!"
    Else
        MsgBox "Sorry, Wrong Door!", vbCritical, "Try Again!"
    End If
End Sub

