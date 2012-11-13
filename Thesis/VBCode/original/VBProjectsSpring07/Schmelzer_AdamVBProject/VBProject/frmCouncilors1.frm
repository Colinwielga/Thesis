VERSION 5.00
Begin VB.Form frmCouncilors1 
   Caption         =   "Real Politik"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSacrifice 
      Caption         =   "We all must sacrifice for our people and land."
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
      Left            =   6840
      TabIndex        =   4
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton cmdLands 
      Caption         =   "Upon victory new lands shall be awarded to you."
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
      Left            =   4560
      TabIndex        =   3
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton cmdPlunder 
      Caption         =   "You may keep whatever you plunder in the war to come."
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
      Left            =   6840
      TabIndex        =   2
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton cmdIlead 
      BackColor       =   &H00C0E0FF&
      Caption         =   "I shall lead and you shall follow."
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
      Left            =   4560
      TabIndex        =   1
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H008080FF&
      Caption         =   $"frmCouncilors1.frx":0000
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   4560
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   0
      Picture         =   "frmCouncilors1.frx":015D
      Top             =   0
      Width           =   9480
   End
End
Attribute VB_Name = "frmCouncilors1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'to give the user different options through command buttons that deal with the
'fictional high council in the game
'as usual, the choices affect some of the main variables (resources, battlepoints,
'econpoints, elerderpoints, and introduced here: politikpoints)

Private Sub cmdIlead_Click()
Politikpoints = Politikpoints - 2
'the men of the army respect you more
Battlepoints = Battlepoints + 100
frmCouncilors1.Hide
frmArmy2.Show
End Sub

Private Sub cmdLands_Click()
Politikpoints = Politikpoints + 1
frmCouncilors1.Hide
frmArmy2.Show
End Sub

Private Sub cmdPlunder_Click()
Politikpoints = Politikpoints + 2
'user will lose resources with this option
Resources = Resources - 100
frmCouncilors1.Hide
frmArmy2.Show
End Sub

Private Sub cmdSacrifice_Click()
Politikpoints = Politikpoints - 1
'the people respect your equal treating of the nobles
Econpoints = Econpoints + 1
frmCouncilors1.Hide
frmArmy2.Show
End Sub


