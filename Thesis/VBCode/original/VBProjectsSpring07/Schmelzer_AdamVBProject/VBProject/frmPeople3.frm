VERSION 5.00
Begin VB.Form frmPeople3 
   BackColor       =   &H00000040&
   Caption         =   "The People"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   Picture         =   "frmPeople3.frx":0000
   ScaleHeight     =   9165
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContract 
      BackColor       =   &H00808080&
      Caption         =   "Tell them that we will draw up new (lucrative) contracts for them to supply the army."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton cmdSpeak 
      BackColor       =   &H00808080&
      Caption         =   "I will speak to them as countrymen and illustrate the need for us all to sacrifice."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton cmdNothing 
      BackColor       =   &H00808080&
      Caption         =   "Let them bicker.       I will do nothing."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton cmdTaxes 
      BackColor       =   &H00808080&
      Caption         =   "Tell them that the commerce tax will be 2% less for the next year if they cease this nonsense."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00000040&
      Caption         =   $"frmPeople3.frx":17AB9
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   6480
      Width           =   4935
   End
End
Attribute VB_Name = "frmPeople3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form presents the user with the several options of command buttons, which
'corresond to the situation desscribed in the form's label box
'choosing a specific command button may affect the resources, econpoints
'also, for the cmdSpeak the boolean variable "orator" dictates the degree to which
'the aforementioned variables are affected
'whatever decision, the present form is hidden and another form made visible


Private Sub cmdContract_Click()
Resources = Resources + 100
Econpoints = Econpoints + 2
If LannisterAllianceP = True Then
    frmPeople3.Hide
    frmAlliance4.Show
End If
If LannisterAllianceP = False And LannisterAllianceN = True Then
    frmPeople3.Hide
    frmArmy3.Show
End If
If LannisterAllianceP = False And LannisterAllianceN = False Then
    frmPeople3.Hide
    frmArmy3noalliance.Show
End If
End Sub

Private Sub cmdNothing_Click()
Econpoints = Econpoints - 2
If LannisterAllianceP = True Then
    frmPeople3.Hide
    frmAlliance4.Show
End If
If LannisterAllianceP = False And LannisterAllianceN = True Then
    frmPeople3.Hide
    frmArmy3.Show
End If
If LannisterAllianceP = False And LannisterAllianceN = False Then
    frmPeople3.Hide
    frmArmy3noalliance.Show
End If
End Sub

Private Sub cmdSpeak_Click()
If Orator = True Or (Courage = True And Scholar = True) Then
    Econpoints = Econpoints + 1
    Resources = Resources + 100
Else
    Econpoints = Econpoints - 1
End If
If LannisterAllianceP = True Then
    frmPeople3.Hide
    frmAlliance4.Show
End If
If LannisterAllianceP = False And LannisterAllianceN = True Then
    frmPeople3.Hide
    frmArmy3.Show
End If
If LannisterAllianceP = False And LannisterAllianceN = False Then
    frmPeople3.Hide
    frmArmy3noalliance.Show
End If
End Sub

Private Sub cmdTaxes_Click()
Resources = Resources - 100
Econpoints = Econpoints - 0
If LannisterAllianceP = True Then
    frmPeople3.Hide
    frmAlliance4.Show
End If
If LannisterAllianceP = False And LannisterAllianceN = True Then
    frmPeople3.Hide
    frmArmy3.Show
End If
If LannisterAllianceP = False And LannisterAllianceN = False Then
    frmPeople3.Hide
    frmArmy3noalliance.Show
End If
End Sub


