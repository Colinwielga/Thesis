VERSION 5.00
Begin VB.Form frm1on1 
   BackColor       =   &H00000080&
   Caption         =   "Truth by Combat"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   Picture         =   "frm1on1.frx":0000
   ScaleHeight     =   9315
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Capitulate"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9840
      TabIndex        =   4
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton cmdSurrender 
      Caption         =   "I shall save the men's lives as well as my own and hopefully my title and some of my wealth and lands."
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
      Left            =   9840
      TabIndex        =   3
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdTowar 
      Caption         =   "I shall rally the men.  Glory is the path this day presents.  The man will heed my call!"
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
      Left            =   9840
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdTrialbycombat 
      Caption         =   "I shall challenge Bolten to single combat and decide this contest once and for all..."
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
      Left            =   9840
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H0080FFFF&
      Caption         =   $"frm1on1.frx":F982
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
      Left            =   0
      TabIndex        =   0
      Top             =   7560
      Width           =   9735
   End
End
Attribute VB_Name = "frm1on1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form presents the user with the option to either fight one on one, surrender,
'or attempt to rally his army for a final battle
'success or the level of success with each of these options depends on
'the values of certain varibales depending on the option taken.
'the user will lose the individual battle if his skilk and personality boolean
'variables to not match certain combinations
'winning and losing present different forms and with these forms, a variety of message
'boxes present different outcomes within the general outcome shown by the form subsequently
'displayed
'the user will go to the final battle regardless of whether or not he holds certain
'values of variables.  Nonetheless, if he has one of three variables his battlepoints
'will increase.  Otherwise they will decrease
'If the user chooses individual battle the main variable affected is the 'victor'
'boolean variable which dictates the subsequent outcomes which are then also dependent
'on the 'eldervariable' (which is true is the user did not successfully kill the elders
'in a previous form)
'the eldervariable dictates whether or not the users total point sum of politikpoints,
'econpoints, and elder points included the elderpoints or not.
'If the elders are dead then the user has to meet a point sum over 0 discluding the
'elderpoints
'whereas if the elders still live their points are included in the sum which is compared
'to a higher value
'Whether or not the user point sum total is greater than a certain amount, which is
'dependent upon the eldervariable dictates whether or not the user succeeds, or succeeds
'in gaining a certain suboutcome of the general outcome
'the suboutcome is conveyed via a msgbox, and the general outcome with a new form



Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdSurrender_Click()
frm1on1.Hide
frmSubmission.Show
End Sub

Private Sub cmdTowar_Click()
If Orator = True Or Courage = True And Scholar = True Then
    Battlepoints = Battlepoints + 1000
    frm1on1.Hide
    frmFinalbattle.Show
Else
    Battlepoints = Battlepoints - 1000
    frm1on1.Hide
    frmFinalbattle.Show
End If
End Sub

Private Sub cmdTrialbycombat_Click()
victor = False
If Courage = True And Strength = True Then
    victor = True
    frm1on1.Hide
    frmLonevictor.Show
    'for victory with bolten submission to user or victory with
    'lannister alliances or tainted victory  or one on one victor only
    If eldervariable = True Then
        If (Econpoints + Politikpoints + Elderpoints) > 4 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. With this victory and your reputation with the poeple, elders, and your high councilors, your legacy has been attained and your honor affirmed.", , "Success."
        End If
        If (Econpoints + Politikpoints + Elderpoints) < 4 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. While you have proved victorious your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
    If eldervariable = False Then
        If (Econpoints + Politikpoints) > 0 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. With this victory and your reputation with the poeple, elders, and your high councilors, your legacy has been attained and your honor affirmed.", , "Success."
        End If
        If (Econpoints + Politikpoints) < 0 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. While you have proved victorious your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
End If
If Courage = True And Warrior = True Then
    victor = True
    frm1on1.Hide
    frmLonevictor.Show
    'for victory with bolten submission to user or victory with
    'lannister alliances or tainted victory  or one on one victor only
    If eldervariable = True Then
        If (Econpoints + Politikpoints + Elderpoints) > 4 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. With this victory and your reputation with the poeple, elders, and your high councilors, your legacy has been attained and your honor affirmed.", , "Success."
        End If
        If (Econpoints + Politikpoints + Elderpoints) < 4 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. While you have proved victorious your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
    If eldervariable = False Then
        If (Econpoints + Politikpoints) > 0 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. With this victory and your reputation with the poeple, elders, and your high councilors, your legacy has been attained and your honor affirmed.", , "Success."
        End If
        If (Econpoints + Politikpoints) < 0 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. While you have proved victorious your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
End If
If Strength = True And Warrior = True Then
    victor = True
    frm1on1.Hide
    frmLonevictor.Show
    'for victory with bolten submission to user or victory with
    'lannister alliances or tainted victory  or one on one victor only
    If eldervariable = True Then
        If (Econpoints + Politikpoints + Elderpoints) > 4 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. With this victory and your reputation with the poeple, elders, and your high councilors, your legacy has been attained and your honor affirmed.", , "Success."
        End If
        If (Econpoints + Politikpoints + Elderpoints) < 4 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. While you have proved victorious your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
    If eldervariable = False Then
        If (Econpoints + Politikpoints) > 0 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. With this victory and your reputation with the poeple, elders, and your high councilors, your legacy has been attained and your honor affirmed.", , "Success."
        End If
        If (Econpoints + Politikpoints) < 0 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. While you have proved victorious your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
End If
If Cunning = True And Warrior = True Then
    victor = True
    frm1on1.Hide
    frmLonevictor.Show
    'for victory with bolten submission to user or victory with
    'lannister alliances or tainted victory  or one on one victor only
    If eldervariable = True Then
        If (Econpoints + Politikpoints + Elderpoints) > 4 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. With this victory and your reputation with the poeple, elders, and your high councilors, your legacy has been attained and your honor affirmed.", , "Success."
        End If
        If (Econpoints + Politikpoints + Elderpoints) < 4 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. While you have proved victorious your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
    If eldervariable = False Then
        If (Econpoints + Politikpoints) > 0 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. With this victory and your reputation with the poeple, elders, and your high councilors, your legacy has been attained and your honor affirmed.", , "Success."
        End If
        If (Econpoints + Politikpoints) < 0 Then
            MsgBox "The men shout at your triumph over the sinister Bolten.  In your defiance of his would-be tyranny your deftness in battle proclaimed your glory as you smote your enemies ruin on the battlefield with ribbons of blood. While you have proved victorious your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
End If
If victor = False Then
    frm1on1.Hide
    frmDeathbed.Show
End If
End Sub

