VERSION 5.00
Begin VB.Form frmFinalbattle 
   BackColor       =   &H00404040&
   Caption         =   "The Final Battle"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9045
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   Picture         =   "frmFinalbattle.frx":0000
   ScaleHeight     =   8025
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFight 
      BackColor       =   &H00C0C0C0&
      Caption         =   "So be it."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   3735
   End
   Begin VB.CommandButton cmdSurrender 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmFinalbattle.frx":118B0
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   3735
   End
   Begin VB.Label lblInstructions2 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmFinalbattle.frx":1193A
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   3
      Top             =   4680
      Width           =   5295
   End
   Begin VB.Label lblInstructions1 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmFinalbattle.frx":11A21
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   4680
      Width           =   5295
   End
End
Attribute VB_Name = "frmFinalbattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form presents the user with one of two options: to fight or to surrender
'surrendering presents a final outcome form, which gives the user feedback as to his
'failing the game
'if the user choses to fight then the victory boolean variable is made true or false
'if the users battlepoints (size of the army) is greater than the enemies, the victory
'variable becomes true, if not it remains false
'if the victory variable is false then one of two failure outcome forms are presented
'dependent upon whether the user chose the boolean variable 'lover' or not
'if the victory variable is true then the ultimate outcome other than success in this battle
'becomes dependent upon his alliance variables, and econ, politik, elderpoints variables,
'and the eldervariable (boolean)
'the eldervariable dictates whether or not the users total point sum of politikpoints,
'econpoints, and elder points included the elderpoints or not.
'If the elders are dead then the user has to meet a point sum over 0 discluding the
'elderpoints
'whereas if the elders still live their points are included in the sum which is compared
'to a higher value
'Whether or not the user point sum total is greater than a certain amount, which is
'dependent upon the eldervariable, dictates whether or not the user succeeds, or succeeds
'in gaining a certain suboutcome of the general outcome
'the suboutcome is conveyed with either a a msgbox, new form, or both

Private Sub cmdFight_Click()
Dim victory As Boolean
victory = False
MsgBox "Boltens forces: " & BoltenArmy & " Your forces: " & Battlepoints, , "Army Sizes"
MsgBox "Your availables resources: " & Resources, , "Resources before last troop mustering"
MsgBox "Elder points: " & Elderpoints & " Polical Points: " & Politikpoints & " Economy Points: " & Econpoints & " Total Points: " & (Econpoints + Politikpoints + Elderpoints) & "Total Points if elders dead: " & (Politikpoints + Econpoints), , "Point Layout"
If Battlepoints > BoltenArmy Then
    victory = True
End If
If victory = False And Lover = False Then
    frmFinalbattle.Hide
    frmFailure.Show
End If
If victory = False And Lover = True Then
    frmFinalbattle.Hide
    frmDeathbutheir.Show
End If
If victory = True And LannisterAllianceP = True Then
    frmFinalbattle.Hide
    frmVictorfornow.Show
    'for victory with bolten submission to user or victory with
    'lannister alliances or tainted victory  or one on one victor only
    If eldervariable = True Then
        If (Econpoints + Politikpoints + Elderpoints) > 4 Then
            MsgBox "You have proved to be a successful leader.  With this victory and your reputation with the people, elders, and your high councilors, your legacy has been attained and your honor affirmed, though it is not yours alone as you allied yourself with Lannister.", , "Success."
        End If
        If (Econpoints + Politikpoints + Elderpoints) < 4 Then
            MsgBox "While you have proved victorious in battle your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
    If eldervariable = False Then
        If (Econpoints + Politikpoints) > 0 Then
            MsgBox "You have proved to be a successful leader.  With this victory and your reputation with the people, elders, and your high councilors, your legacy has been attained and your honor affirmed, though it is not yours alone as you allied yourself with Lannister, though it is not yours alone as you allied yourself with Lannister.", , "Success."
        End If
        If (Econpoints + Politikpoints) < 0 Then
            MsgBox "While you have proved victorious in battle your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
End If
If victory = True And LannisterAllianceN = True And LannisterLifeV = True Then
    frmFinalbattle.Hide
    frmVictorfornow.Show
    'for victory with bolten submission to user or victory with
    'lannister alliances or tainted victory  or one on one victor only
    If eldervariable = True Then
        If (Econpoints + Politikpoints + Elderpoints) > 4 Then
            MsgBox "Though set back at first, you have proved to be a somewhat successful leader.  With this victory and your reputation with the people, elders, and your high councilors, your legacy has become a possibility.  Nevertheless, while you live a Lord, Lannister is still your superior.  Your family honor has been protected but only at a loss of its worth.", , "Bitter Success."
        End If
        If (Econpoints + Politikpoints + Elderpoints) < 4 Then
            MsgBox "While you have proved victorious in battle your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
    If eldervariable = False Then
        If (Econpoints + Politikpoints) > 0 Then
            MsgBox "Though set back at first, you have proved to be a somewhat sucsessful leader.  With this victory and your reputation with the people, elders, and your high councilors, your legacy has become a possibility.  Nevertheless, while you live a Lord, Lannister is still your superior.  Your family honor has been protected but only at a loss of its worth.", , "Bitter Success."
        End If
        If (Econpoints + Politikpoints) < 0 Then
            MsgBox "While you have proved victorious in battle your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
End If
If victory = True And LannisterAllianceN = True And LannisterLifeV = False Then
    frmFinalbattle.Hide
    frmVictorybykill.Show
    'for victory with bolten submission to user or victory with
    'lannister alliances or tainted victory  or one on one victor only
    If eldervariable = True Then
        If (Econpoints + Politikpoints + Elderpoints) > 4 Then
            MsgBox "Though set back at first, you have proved to be a cunning and sucsessful leader.  With this victory and your reputation with the people, elders, and your high councilors, your legacy has been attained and your honor ironically and unknowingly affirmed.  Nevertheless, while uour family honor has been protected, it has been at a loss of its worth.", , "Victor by Blood."
        End If
        If (Econpoints + Politikpoints + Elderpoints) < 4 Then
            MsgBox "While you have proved victorious in battle and in your assassination of Lannister, your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
    If eldervariable = False Then
        If (Econpoints + Politikpoints) > 0 Then
            MsgBox "Though set back at first, you have proved to be a cunning and sucsessful leader.  With this victory and your reputation with the people, elders, and your high councilors, your legacy has been attained and your honor ironically and unknowingly affirmed.  Nevertheless, while uour family honor has been protected, it has been at a loss of its worth.", , "Victor by Blood."
        End If
        If (Econpoints + Politikpoints) < 0 Then
            MsgBox "While you have proved victorious in battle and in your assassination of Lannister, your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
End If
If victory = True And LannisterAllianceP = False And LannisterAllianceN = False Then
     If eldervariable = True Then
        If (Econpoints + Politikpoints + Elderpoints) > 4 Then
            frmFinalbattle.Hide
            frmGlory.Show
        End If
        If (Econpoints + Politikpoints + Elderpoints) < 4 Then
            frmFinalbattle.Hide
            frmVictorfornow.Show
            MsgBox "While you have proved victorious in battle and in your assassination of Lannister, your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
    If eldervariable = False Then
        If (Econpoints + Politikpoints) > 0 Then
            frmFinalbattle.Hide
            frmGlory.Show
        End If
        If (Econpoints + Politikpoints) < 0 Then
            frmFinalbattle.Hide
            frmVictorfornow.Show
            MsgBox "While you have proved victorious in battle and in your assassination of Lannister, your failings in ruling the realm ultimately led to your heir being murdered and his throne usurped by either Bolten or Lannister.  Either your high councilors and kingdom's noble families or the priesthood or both will have conspired against your house and left the memory of your family to be lost in the ashes of history.  Though you attained victory in this instance, the ultimate victory--that of legacy and longevity--has been lost.", , "You have failed."
        End If
    End If
End If
    
End Sub

Private Sub cmdSurrender_Click()
frmFinalbattle.Hide
frmSubmission.Show
End Sub

Private Sub Form_Load()
If LannisterAllianceP = True Or LannisterAllianceN = True Then
    lblInstructions1.Visible = False
    lblInstructions2.Visible = True
End If
If LannisterAllianceP = False And LannisterAllianceN = False Then
    lblInstructions1.Visible = True
    lblInstructions2.Visible = False
Else
    lblInstructions1.Visible = True
    lblInstructions2.Visible = False
End If
End Sub


