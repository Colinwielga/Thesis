VERSION 5.00
Begin VB.Form frmElders22 
   Caption         =   "The Elders"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   Picture         =   "frmElders22.frx":0000
   ScaleHeight     =   9720
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPayoff 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Let's bribe the elders.  Pay them off.  5,000 gold crowns to the faith.  "
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton cmdDefy 
      BackColor       =   &H00C0FFFF&
      Caption         =   "We will fight with mine own banners and that of the army. (Defy them)."
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdCarry 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Trust the gods.  I will carry their banner."
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFC0&
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
      Height          =   1095
      Left            =   7200
      MaskColor       =   &H0080FFFF&
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmElders22.frx":1A0FB
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmElders22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form presents the user with three options as command buttons that relate
'to the situation described in the form's label
'depending on which option is chosen different variables are affected
'these include elderpoints, battlepoints, and politikpoints
'Furthermore, if the option to bribe the elders is chosen
'the users success is dependent upon his procurement of certain combinations
'of the skill and personality boolean variables which were chosen via the first form
'and input boxes as public boolean variables
'if the combinational requirements are met then the user receives positive points and feedback,
'respectively.  Otherwise the opposite occurs
'A new for is made visible


Private Sub cmdCarry_Click()
Elderpoints = Elderpoints + 2
Battlepoints = Battlepoints - 100
Politikpoints = Politikpoints - 1
MsgBox "The elders are pleased with their new puppet.", , "Praise the gods"
frmElders22.Hide
frmPeople2.Show
End Sub

Private Sub cmdDefy_Click()
Elderpoints = Elderpoints - 1
Battlepoints = Battlepoints + 100
Politikpoints = Politikpoints + 1
MsgBox "The elders will not easily forget this.  Yet the men and the council rally behind your bravado.", , "Strength and Honor"
frmElders22.Hide
frmPeople2.Show
End Sub

Private Sub cmdPayoff_Click()
Resources = Resources - 50
If (Cunning = True And Strength = True) Or (Scholar = True And Orator = True) Or (Courage = True And Orator = True) Or (Cunning = True And Orator = True) Or (Strength = True And Scholar = True) Or (Cunning = True And Scholar = True) Or (Courage = True And Cunning = True) Then
    Elderpoints = Elderpoints + 2
    MsgBox "Your cunning and strength of resovle won out.  The elders' piety was purchased. It is yours, for now.", , "A capital purchase"
    frmElders22.Hide
    frmPeople2.Show
Else
    Elderpoints = Elderpoints - 2
    MsgBox "Due to your lack of cunning, wit, and political skill, the elders refused your offer.  You now play the fool, to the council, army, and people.", , "Failing Investment"
    frmElders22.Hide
    frmPeople2.Show
End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

