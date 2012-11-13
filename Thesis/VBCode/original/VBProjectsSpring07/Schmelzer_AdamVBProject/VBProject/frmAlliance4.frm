VERSION 5.00
Begin VB.Form frmAlliance4 
   BackColor       =   &H00404040&
   Caption         =   "Submission"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   Picture         =   "frmAlliance4.frx":0000
   ScaleHeight     =   9090
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDecline 
      Caption         =   "There shall be battle, and bloodshed enough to warm this cold earth."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   9960
      TabIndex        =   3
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Capitulate"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9960
      TabIndex        =   2
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton cmdBend 
      Caption         =   "I accept his offer."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9960
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00808000&
      Caption         =   $"frmAlliance4.frx":E3D2
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
      Left            =   0
      TabIndex        =   0
      Top             =   7560
      Width           =   9735
   End
End
Attribute VB_Name = "frmAlliance4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'the form present the user with two options: accepting or delcining an alliance,
'which in this case is the users submission and forfeiture of further game play
'if the user declines, the only variable affected is the alliances boolean variable
'if accpeted the boolean variable is changed to true, and a series of nested if statements
'dictates what will happen next
'The nested if statements dictate which sub outcomes will appear via message boxes
'within the general outcome shown in the 'submission' for that is subsequently displayed
'the eldervariable dictates whether or not the user's total point sum of politikpoints,
'econpoints, and elder points included the elderpoints or not.
'If the elders are dead then the user has to meet a point sum over 0 discluding the
'elderpoints
'whereas if the elders still live their points are included in the sum which is compared
'to a higher value
'Whether or not the user point sum total is greater than a certain amount, which is
'dependent upon the eldervariable dictates whether or not the user succeeds, or succeeds
'in gaining a certain suboutcome of the general outcome
'the suboutcome is conveyed via a msgbox, and the general outcome with a new form

Private Sub cmdBend_Click()
BoltenAllianceP = True
frmAlliance4.Hide
frmVictorfornow.Show
'for victory with bolten submission to user or victory with
    'lannister alliances or tainted victory  or one on one victor only
    If eldervariable = True Then
        If (Econpoints + Politikpoints + Elderpoints) > 4 Then
            MsgBox "Though set back at first, you have proved to be a sucessful leader.  With this victory and your reputation with the poeple, elders, and your high councilors, your legacy has been attained and your honor affirmed.", , "Success."
        End If
        If (Econpoints + Politikpoints + Elderpoints) < 4 Then
            MsgBox "While you have proved victorious your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
    If eldervariable = False Then
        If (Econpoints + Politikpoints) > 0 Then
            MsgBox "Though set back at first, you have proved to be a sucessful leader.  With this victory and your reputation with the poeple, elders, and your high councilors, your legacy has been attained and your honor affirmed.", , "Success."
        End If
        If (Econpoints + Politikpoints) < 0 Then
            MsgBox "While you have proved victorious your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
End Sub

Private Sub cmdDecline_Click()
frmAlliance4.Hide
frmArmy3.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub
