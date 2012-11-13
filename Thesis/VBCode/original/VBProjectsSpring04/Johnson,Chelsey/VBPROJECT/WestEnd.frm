VERSION 5.00
Begin VB.Form WestEnd 
   BackColor       =   &H00FFFF80&
   Caption         =   "West End"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   14760
      TabIndex        =   10
      Top             =   9840
      Width           =   2175
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Map of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   12000
      TabIndex        =   9
      Top             =   8880
      Width           =   2415
   End
   Begin VB.PictureBox picwaterloo 
      Height          =   6015
      Left            =   10320
      Picture         =   "WestEnd.frx":0000
      ScaleHeight     =   5955
      ScaleWidth      =   8715
      TabIndex        =   7
      Top             =   2160
      Width           =   8775
   End
   Begin VB.PictureBox picTheatre 
      Height          =   3015
      Left            =   3600
      Picture         =   "WestEnd.frx":DE74
      ScaleHeight     =   2955
      ScaleWidth      =   6915
      TabIndex        =   5
      Top             =   8280
      Width           =   6975
   End
   Begin VB.PictureBox picResults 
      Height          =   4095
      Left            =   3720
      ScaleHeight     =   4035
      ScaleWidth      =   6435
      TabIndex        =   4
      Top             =   2280
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Text            =   "***Click on the picture to learn more about this famous site"
      Top             =   1560
      Width           =   5295
   End
   Begin VB.PictureBox picNeedle 
      Height          =   10335
      Left            =   480
      Picture         =   "WestEnd.frx":146C0
      ScaleHeight     =   10275
      ScaleWidth      =   2835
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   12000
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Waterloo Bridge is a very famous site***To learn more about this bridge, click on the picture."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10440
      TabIndex        =   8
      Top             =   960
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackColor       =   &H008080FF&
      Caption         =   "This is the National Theatre*** Click on the picture to learn more about it."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      TabIndex        =   6
      Top             =   7080
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF80FF&
      Caption         =   "*** This is Cleopatra's Needle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "West End District Has Many Facinating Sites"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "WestEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: WestEnd (WestEnd.frm)
'Author: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of Form: This form gives the user the option of clicking on a picture of their choice to learn about the history
                'of that picture
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'User returns to the Map of London page to choose a new district
WestEnd.Hide
MapLondon.Show
End Sub

Private Sub picNeedle_Click()
picResults.Cls
'Prints the history of Cleopatra's Needle
picResults.Print "Cleopatra's Needle was made in Egypt for the Pharaoh Thotmes III in 1460 BC"
picResults.Print "making it almost 3,500 years old."
picResults.Print "It is known as Cleopatra's Needle as it was brought to London"
picResults.Print "from Alexandria, the royal city of Cleopatra. "
picResults.Print "The Needle arrived in England after a horrendous journey by sea in 1878."
picResults.Print "Cleopatra's Needle stands on the Thames Embankment close"
picResults.Print "to the Embankment underground station. "
picResults.Print "Two large bronze Sphinxes lie on either side of the Needle."
picResults.Print "These are Victorian versions of the traditional Egyptian original. "
picResults.Print "The benches on the Embankment also have winged sphinxes "
picResults.Print "on either side as their supports."
End Sub

Private Sub picTheatre_Click()
picResults.Cls
'Prints the history of The National Theatre
picResults.Print "The Royal National Theatre was designed by Dennis Lasdun."
picResults.Print "There are three auditoria-"
picResults.Print "The Olivier, Lyttleton and Cottesloe."
picResults.Print "The National Theatre Company was founded in 1962"
picResults.Print "originally based at the Old Vic"
picResults.Print "where they played with great success for nearly 13 years."
End Sub

Private Sub picwaterloo_Click()
picResults.Cls
'Prints the history of The Waterloo Bridge
picResults.Print "This noble bridge is from the designs of John Rennie"
picResults.Print "the distinguished engineer of the Plymouth Breakwater and many other well-known works."
picResults.Print "The first stone of this celebrated bridge was laid on the 11th October, 1811"
picResults.Print "and the bridge was completed in less than six years, the public opening"
picResults.Print "taking place on the second anniversary of the battle of Waterloo, June 18th, 1817."
picResults.Print "The total cost was over one million of money."
picResults.Print "The design differs from all other bridges over the Thames, being one uniform "
picResults.Print "level throughout its entire length of 2,456 feet."
picResults.Print "Its construction consists of nine elliptical arches of 120 feet span"
picResults.Print "and 35 feet high, supported on piers 20 feet wide at the springing of the arches."
End Sub
