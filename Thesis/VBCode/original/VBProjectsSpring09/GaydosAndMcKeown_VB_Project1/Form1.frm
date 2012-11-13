VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   -135
   ClientTop       =   2205
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8040
   ScaleWidth      =   8715
   Begin VB.CommandButton Command1 
      Caption         =   "Surrender"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   2160
      Picture         =   "Form1.frx":55D02
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   4335
   End
   Begin VB.CommandButton cmdWarsaw 
      Caption         =   "Warsaw Pact"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4320
      Picture         =   "Form1.frx":5AD53
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   4575
   End
   Begin VB.CommandButton cmdNato 
      Caption         =   "NATO"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      Picture         =   "Form1.frx":5D68A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "SELECT WHICH FORCES AMMUNITION YOU WANT TO LOOK AT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "COLD WAR AMMUNITION PROJECT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNato_Click()
Form1.Hide 'this effectively ends the main menu
frmNATO.Show 'this button makes the nato form show
'This button goes to the NATO Form


End Sub

Private Sub cmdWarsaw_Click()
Form1.Hide 'this effectively ends the main menu
frmWarsaw.Show 'this button makes the Warsaw Pact show
'This button shows the Warsaw Pact Form
End Sub

Private Sub Command1_Click()
MsgBox "What??? YOU COULDN'T HANDLE IT?"
End

End Sub

'You can exit from this program at the start!

'We chose to have a main menu so that we could fit the two faction's munitions on different forms
'This saves space, as well as gives us opportunities to try out different codes for each form
'Since we have an end program and switch form on both of the other forms, we decided there was no need to go back to the main menu, so the first time it has been used is the last time it will be used until the next time the user opens this project
