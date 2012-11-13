VERSION 5.00
Begin VB.Form frmAssassination 
   Caption         =   "Guile."
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   Picture         =   "frmAssassination.frx":0000
   ScaleHeight     =   8745
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDeny 
      BackColor       =   &H00C0FFFF&
      Caption         =   "No, he is both friend and ally."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton cmdKill 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Do it. Kill him."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAssassination.frx":26B5F
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmAssassination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form presents the user with the options of either attempting to assassinate
'his ally, Lord Lannister, or declining to do so.
'if the user declines nothing more than the switching of visible forms occurs
'if the user chooses the command button of attempting an assassination,
'the succuss of such is dependent upon the value of his politikpoints (single)
'AND whether or not he possess a certain combination of 'skill' and 'personality'
'variables, where possessing means that such variables are true
'if he succeeds or fails one of two forms makes itself visible to the user
'and either brings him to the final battle form or takes him to his now last
'form and outcome of death, from which the program ends


Private Sub cmdDeny_Click()
    frmAssassination.Hide
    frmFinalbattle.Show
End Sub

Private Sub cmdKill_Click()
Dim Success As Boolean
Success = False
If Politikpoints > 3 And Cunning = True Then
    MsgBox "With your influence in political circles, your wit, and cunning, you will pull off a masterstroke in tomorrows battle.  Nonetheless, the outcome of the war is still to be decided.  Good luck!", , "A Masterstroke"
    LannisterLifeV = False
    frmAssassination.Hide
    frmFinalbattle.Show
    Success = True
End If
If Politikpoints > 3 And Courage = True And Scholar = True Then
    MsgBox "With your influence in political circles, your wit, and cunning, you will pull off a masterstroke in tomorrows battle.  Nonetheless, the outcome of the war is still to be decided.  Good luck!", , "A Masterstroke"
    LannisterLifeV = False
    frmAssassination.Hide
    frmFinalbattle.Show
    Success = True
End If
If Politikpoints > 3 And Courage = True And Orator = True Then
    MsgBox "With your influence in political circles, your wit, and cunning, you will pull off a masterstroke in tomorrows battle.  Nonetheless, the outcome of the war is still to be decided.  Good luck!", , "A Masterstroke"
    LannisterLifeV = False
    frmAssassination.Hide
    frmFinalbattle.Show
    Success = True
End If
If Politikpoints > 3 And Scholar = True And Orator = True Then
    MsgBox "With your influence in political circles, your wit, and cunning, you will pull off a masterstroke in tomorrows battle.  Nonetheless, the outcome of the war is still to be decided.  Good luck!", , "A Masterstroke"
    LannisterLifeV = False
    frmAssassination.Hide
    frmFinalbattle.Show
    Success = True
    
End If
If Success = False Then
    LannisterLifeV = True
    frmAssassination.Hide
    frmAssassinationfailure.Show
End If
End Sub
