VERSION 5.00
Begin VB.Form frmArmy2 
   BackColor       =   &H00000080&
   Caption         =   "Army II"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdField 
      Caption         =   "I shall wait for him on the field.  Let him be man enought to take it."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4920
      TabIndex        =   5
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Capitulate"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7200
      TabIndex        =   4
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdSiege 
      Caption         =   "Ready the Siege!"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2520
      TabIndex        =   3
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdBlockade 
      Caption         =   "We shall starve him out.  Ready the men! (Blockade)"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   6735
      Left            =   0
      Picture         =   "frmArmy2.frx":0000
      ScaleHeight     =   6675
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   -120
      Width           =   9015
      Begin VB.Label lblInstructions 
         BackColor       =   &H000000C0&
         Caption         =   $"frmArmy2.frx":15C91
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
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   8895
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   15
         Left            =   0
         TabIndex        =   1
         Top             =   6720
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmArmy2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form present the user with different command button options, which relate
'to the situation decribed in the forms label box
'depending on which option the user takes one of several boolean variables changed to true
'those being: waitedV, successfulsiegeV, failedsiegeV, blockadeV
'depending on which command button is chosen; which variable that turns to true and
'how other variables are affected (battlepoints and resources) depends on what type
'of weapons the user invested in, in frmArmy1


Private Sub cmdBlockade_Click()
If siege > 0 Then
    MsgBox "The entirety of the army is calling you a fool; a jester more than king they say.  We have stood outside the castle's fortifications for weeks without action.  We waste time and energy, and let our SIEGE weapons go unused!  I myself am starting to question you my lord.  Nevertheless, despite your recent failing, we must now decide on a new course of action as Bolton's reinforcments (in fact a very large force) are just days away. They are led by his half cousin Lord Gornick who commands the mountain passageways.  These passageways are our only escape route.  You must take to the field, and pray.", , "Danger Looms."
    blockadeV = True
    Battlepoints = Battlepoints - 100
End If
If siege <= 0 Then
    MsgBox "While the blockade proved marginally beneficial in weakening Lord Bolten's forces, your queen mother has been hanged and Bolten's reinforcements led by his half cousin Lord Gornick come from behind us. They will cut off our only escape route.", , "A Bind."
    blockadeV = True
    Battlepoints = Battlepoints - 100
    BoltenArmy = BoltenArmy - 200
End If
'the blockade used up many resources
Resources = Resources - 200
frmArmy2.Hide
frmElders22.Show
End Sub

Private Sub cmdField_Click()
If siege > 0 Then
    MsgBox "Though some thought the castle ripe for the taking especially with the queen mother hostage, others find prudence in your ways.  Only time will tell.  Lord Bolten has left to meet the remainder of his army and heavy reinforcements courtesy of his half cousin and Lord of the mountain passageways, Lord Gornick.  He brings many maurauders my lord.", , "Ambivalence"
    waitedV = True
End If
If siege <= 0 Then
    MsgBox "As we had no siege weaponry prepared for a siege your decision was prudent, though some say dishonorable in light of the queen mother.  Nevertheless, your most important generals and officers find it a wise choice.  Well done sire.  Lord Bolten has left to meet the remainder of his army and heavy reinforcements courtesy of his half cousin and Lord of the mountain passageways, Lord Gornick.  He brings many maurauders my lord.", , "Prudent"
    waitedV = True
End If
frmArmy2.Hide
frmElders22.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSiege_Click()
If siege > 0 Then
    MsgBox "With your investment in Siege Weapons we have taken the castle, and sadly found the ravaged corpse of your mother.  Lord Bolten has fled to meet the remainder of his army and heavy reinforcements courtesy of his half cousin and Lord of the mountain passageways, Lord Gornick.  He brings many maurauders my lord.", , "You are Victorious!"
    successfulsiegeV = True
    Resources = Resources - 200
    Battlepoints = Battlepoints - 200
    BoltenArmy = BoltenArmy - 750
End If
If siege <= 0 Then
    MsgBox "Foolishly you led the men to a siege with no siege weapons prepared.  Many have died in weak attempts at taking the gate.  As our resources dwindle we are forced to retreat.  Lord Bolten has fled to meet the remainder of his army and heavy reinforcements courtesy of his half cousin and Lord of the mountain passageways, Lord Gornick.  He brings many maurauders my lord.", , "Failure."
    failsiegeV = True
    'in the siege this many units were lost
    Battlepoints = Battlepoints - 500
    Resources = Resources - 300
    BoltenArmy = BoltenArmy - 200
End If
frmArmy2.Hide
frmElders22.Show
End Sub

Private Sub Form_Load()
BoltenArmy = BoltenArmy + 3000
End Sub
