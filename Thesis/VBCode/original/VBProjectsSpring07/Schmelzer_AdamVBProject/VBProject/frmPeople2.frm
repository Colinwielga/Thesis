VERSION 5.00
Begin VB.Form frmPeople2 
   BackColor       =   &H00004080&
   Caption         =   "The People"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   Picture         =   "frmPeople2.frx":0000
   ScaleHeight     =   7635
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFurnaces 
      BackColor       =   &H00808080&
      Caption         =   "Construct many large furnaces to burn the overflow.  It will be both cheap and effective. "
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton cmdDonothing 
      BackColor       =   &H00808080&
      Caption         =   "Do nothing.  Let them live in what they are."
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
      Left            =   8160
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton cmdRemedy 
      BackColor       =   &H00808080&
      Caption         =   "We shall remedy the problem at once!  Though it will cost the treasury, we are obliged."
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton cmdTaxes 
      BackColor       =   &H00808080&
      Caption         =   "Lower the taxes for now.  The people will find that payment enough I expect."
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
      Left            =   5760
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00004080&
      Caption         =   $"frmPeople2.frx":15C23
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   5535
   End
End
Attribute VB_Name = "frmPeople2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form presents the user with the several options of command buttons, which
'corresond to the situation desscribed in the form's label box
'choosing a specific command button may affect the resources, econpoints, elderpoints
'whatever decision, the present form is hidden and another form is made visible
'the form that is made visible is dependent upon if statements which deal
'with the values of the boolean variables that correspond to the outcome of the frmArmy2
'decisions


Private Sub cmdDonothing_Click()
Econpoints = Econpoints - 2
Resources = Resources - 0
If successfulsiegeV = True Then
    frmPeople2.Hide
    frmAlliance3.Show
End If
If waitedV = True Then
    frmPeople2.Hide
    frmAlliance3.Show
End If
If blockadeV = True Then
    frmPeople2.Hide
    frmAlliance.Show
End If
If failsiegeV = True Then
    frmPeople2.Hide
    frmAlliance2.Show
End If

End Sub

Private Sub cmdFurnaces_Click()
Econpoints = Econpoints + 1
Resources = Resources - 100
Elderpoints = Elderpoints - 1
MsgBox "While this seemed a practical remedy, the elders are in an uproar as this is considered a blaspheme to the heavens.  In reality, both the nobles and the priesthood are angry with you as the stench as risen to the windows of their city residences.  What a terrible inconvenience!", , "Shit happens"
If successfulsiegeV = True Then
    frmPeople2.Hide
    frmAlliance3.Show
End If
If waitedV = True Then
    frmPeople2.Hide
    frmAlliance3.Show
End If
If blockadeV = True Then
    frmPeople2.Hide
    frmAlliance.Show
End If
If failsiegeV = True Then
    frmPeople2.Hide
    frmAlliance2.Show
End If
End Sub

Private Sub cmdRemedy_Click()
Econpoints = Econpoints + 2
Resources = Resources - 200
If successfulsiegeV = True Then
    frmPeople2.Hide
    frmAlliance3.Show
End If
If waitedV = True Then
    frmPeople2.Hide
    frmAlliance3.Show
End If
If blockadeV = True Then
    frmPeople2.Hide
    frmAlliance.Show
End If
If failsiegeV = True Then
    frmPeople2.Hide
    frmAlliance2.Show
End If

End Sub

Private Sub cmdTaxes_Click()
Econpoints = Econpoints - 2
Resources = Resources - 100
If successfulsiegeV = True Then
    frmPeople2.Hide
    frmAlliance3.Show
End If
If waitedV = True Then
    frmPeople2.Hide
    frmAlliance3.Show
End If
If blockadeV = True Then
    frmPeople2.Hide
    frmAlliance.Show
End If
If failsiegeV = True Then
    frmPeople2.Hide
    frmAlliance2.Show
End If
End Sub

