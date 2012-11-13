VERSION 5.00
Begin VB.Form frmYard 
   Caption         =   "Your front yard"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   Picture         =   "frmYard.frx":0000
   ScaleHeight     =   8730
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H8000000A&
      Caption         =   "Continue..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5160
      MaskColor       =   &H8000000A&
      TabIndex        =   3
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton cmdMoney 
      Height          =   1335
      Left            =   8520
      Picture         =   "frmYard.frx":14E386
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "I don't know if there's any stores left, but this may help later if i can buy something"
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton cmdBat 
      Height          =   1815
      Left            =   4920
      Picture         =   "frmYard.frx":14F170
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "This could be a useful weapon if you face an alien"
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdHP 
      Height          =   1695
      Left            =   1200
      Picture         =   "frmYard.frx":14F854
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "These may increase your H.P."
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Some Money"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   6
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Baseball Bat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Bandages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   5760
      Width           =   1335
   End
End
Attribute VB_Name = "frmYard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oldMan As String

Private Sub cmdBat_Click()  'pick up a bat and gain 25 attack
    
    Attack = Attack + 25
    MsgBox ("You picked up a baseball bat.  This could be useful if you have to fight an alien."), , ("Baseball bat")
    MsgBox ("Your attack went up 25 points and is at: " & Attack), , ("Attack = " & Attack)
    
    cmdHP.Enabled = False
    cmdMoney.Enabled = False
    cmdBat.Enabled = False
    cmdContinue.Enabled = True
End Sub

Private Sub cmdContinue_Click() 'continue to street.  old man asks you to enter store.  choose yes or no.
    MsgBox ("Okay, you picked up what you could from your yard, now it's time to head into town.  Let's head down College Ave N."), , ("Continuing on..")
    frmYard.Hide
    frmStreet.Show
    MsgBox ("You enter into the street and see an old man with a sign.  It says: 'Welcome to my store.'"), , ("Old Man")
   
    Do Until oldMan = "yes" Or oldMan = "no"
        oldMan = InputBox("Do you want to go and talk with him?  Type 'yes' or 'no'", "Talk with him?")
        If oldMan = "yes" Then
            frmStore.Show
            frmStreet.Hide
            MsgBox ("'Welcome to Hell, sonny,' the old man says.  'You won't make it in this world without some help so buy somethin' from my store, you'll need it!'"), , ("The Old Man's Store")
            MsgBox ("It looks like he has collected some items and is trying to sell them."), , ("Old Man's Store")
        ElseIf oldMan = "no" Then
            MsgBox ("Okay, forget about that guy, lets head down the street more."), , ("No way I'm talking to that guy")
            frmStreet.Hide
            frmStreet1.Show
        Else
            MsgBox ("Enter 'yes' or 'no'!  You're entry was invalid."), , ("Error")
            
        End If
    Loop
End Sub

Private Sub cmdHP_Click()   'pick up bandages and gain 50 h.p.
    
    HP = HP + 50
    MsgBox ("You picked up bandages.  These will be useful if you get hurt!"), , ("Bandages")
    MsgBox ("Your H.P. went up 50 and is at " & HP), , ("H.P. = " & HP)
    
    cmdHP.Enabled = False
    cmdMoney.Enabled = False
    cmdBat.Enabled = False
    cmdContinue.Enabled = True
End Sub

Private Sub cmdMoney_Click()    'pick up money and gain $150

    Money = Money + 150
    MsgBox ("So, you're the greedy type and want more money.  Well, there aren't any stores still standing, but maybe you could buy something later somehow."), , ("$150")
    MsgBox ("Your money went up $150 and is at " & FormatCurrency(Money)), , ("Money = " & FormatCurrency(Money))
    
    cmdHP.Enabled = False
    cmdMoney.Enabled = False
    cmdBat.Enabled = False
    cmdContinue.Enabled = True
End Sub
