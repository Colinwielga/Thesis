VERSION 5.00
Begin VB.Form frmCouncilors2 
   BackColor       =   &H00404040&
   Caption         =   "The High Council"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   Picture         =   "frmCouncilors2.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCompromise 
      BackColor       =   &H00C0C0C0&
      Caption         =   "You must decide from amongst yourselves"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8520
      Width           =   1815
   End
   Begin VB.CommandButton cmdPay 
      BackColor       =   &H00C0C0C0&
      Caption         =   "If you wish to lead you must purchase the honor."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8520
      Width           =   1935
   End
   Begin VB.CommandButton cmdFight1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Whomever wishes to lead shall fight he whom I appoint."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmdLouis 
      BackColor       =   &H00C0C0C0&
      Caption         =   "I shall make the choice."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"frmCouncilors2.frx":1AA93
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   7200
      Width           =   5295
   End
End
Attribute VB_Name = "frmCouncilors2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'objective: to give the user different options through command buttons that deal
'with the fictional high council in the game
'as usual, the choices affect some of the main variables (resources, battlepoints,
'econpoints, elerderpoints, and introduced here: politikpoints)
'and lead to a new form being visible


Private Sub cmdCompromise_Click()
'this is the equivalent of doing nothing
Politikpoints = Politikpoints - 2
'As different units of soldiers are loyal to different commanders the army has less cohesion
Battlepoints = Battlepoints - 100
frmCouncilors2.Hide
frmElders3.Show
End Sub

Private Sub cmdFight1_Click()
Politikpoints = Politikpoints + 1
'As different units of soldiers are loyal to different commanders the army has less cohesion
Battlepoints = Battlepoints - 100
frmCouncilors2.Hide
frmElders3.Show
End Sub

Private Sub cmdLouis_Click()
Politikpoints = Politikpoints - 0
'the men respect  you more and the commanders fear you, thus the army will run more smoothly assumig user makes wise decisions
Battlepoints = Battlepoints + 100
frmCouncilors2.Hide
frmElders3.Show
End Sub

Private Sub cmdPay_Click()
Politikpoints = Politikpoints + 2
Resources = Resources + 100
'the men respect user's level of integrity less as you base honors not on merit
Battlepoints = Battlepoints - 200
frmCouncilors2.Hide
frmElders3.Show
End Sub
