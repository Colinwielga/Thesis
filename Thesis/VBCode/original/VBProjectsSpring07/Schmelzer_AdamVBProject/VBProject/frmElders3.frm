VERSION 5.00
Begin VB.Form frmElders3 
   BackColor       =   &H000040C0&
   Caption         =   "The Elders"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Capitulate"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8160
      TabIndex        =   5
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton cmdNothing 
      Caption         =   "Do Nothing."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8160
      TabIndex        =   4
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdAgift 
      Caption         =   "Give the women to the old men.  Let them do what they will."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8160
      TabIndex        =   3
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdAdhere 
      Caption         =   "Do as they request.  Kill the women."
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
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdPoison 
      Caption         =   "Poison the Elders, kill them all!  But subtley..."
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
      Left            =   8160
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H000040C0&
      Caption         =   $"frmElders3.frx":0000
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   8055
   End
   Begin VB.Image Image1 
      Height          =   5400
      Left            =   0
      Picture         =   "frmElders3.frx":01B8
      Top             =   0
      Width           =   8100
   End
End
Attribute VB_Name = "frmElders3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'user is presented with a variety of options as commandbuttons that correspond
'to the situation described in the forms label
'the choice of a certain command button affects the elderpoints variable
'if the option of bribing the elders is chosen, the user must meet certain
'requirements, which include having a a sum of elder + politik points > 1 and having
'a one of several combinations of skill and personality boolean variables
'depening on his success or failure, appropriate feedback is displayed via a msgbox
'and the elderpoints single variable is affected appropriately
'a new form is made visible
Private Sub Picture1_Click()

End Sub

Private Sub cmdAdhere_Click()
Elderpoints = Elderpoints + 2
MsgBox "the Elders thirst for blood is sated, but they may be quick to flex their muscle in the future now.", , "Sated"
frmElders3.Hide
frmPeople3.Show
End Sub

Private Sub cmdAgift_Click()
Elderpoints = Elderpoints + 1
MsgBox "An insult to the elders no doubt, but they cannot say a word.  And who knows, they may even enjoy themselves...", , "Swift"
frmElders3.Hide
frmPeople3.Show
End Sub

Private Sub cmdNothing_Click()
Elderpoints = Elderpoints - 1
MsgBox "The elders are furious, but confused.  They do not know what to do in response to  your silence.  Yet, they are rallying the faithful masses and are seeking to the reconception of the army of the Gods.  A religious militia may be seen in the streets soon.", , "Beware"
frmElders3.Hide
frmPeople3.Show
End Sub

Private Sub cmdPoison_Click()
If (Elderpoints + Politikpoints) > 1 And Cunning = True Then
    MsgBox "A masterstroke sire.  With your cunning and built up capital with the high council and elders you were able to rid yourself of those fanatics once and for all.", , "Success"
    eldervariable = False
Else
    MsgBox "Your mistresses were seized, tortured to 'purity', and finally after several days killed.  With your lack of cunning and and your disfavor in certain circles, your attempt was expected.  The elders now call for your blood.  Be wary, my king, extremely so...", , "Failure."
    Elderpoints = -2
    'lowest points possible
End If
frmElders3.Hide
frmPeople3.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub
