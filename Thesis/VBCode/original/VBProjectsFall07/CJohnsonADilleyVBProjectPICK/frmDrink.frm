VERSION 5.00
Begin VB.Form frmDrink 
   BackColor       =   &H00FFFF00&
   Caption         =   "Drinking Buddy!"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15450
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H00FFFF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   15450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBegin 
      BackColor       =   &H008080FF&
      Caption         =   "Click Here To Start Your Night!"
      BeginProperty Font 
         Name            =   "Charlemagne Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Width           =   4695
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFF00&
      Caption         =   $"frmDrink.frx":0000
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   13695
   End
   Begin VB.Label lblBuddy 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your Drinking Buddy"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   0
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Image partyimage 
      Height          =   4500
      Left            =   4800
      Picture         =   "frmDrink.frx":029C
      Top             =   3240
      Width           =   6000
   End
End
Attribute VB_Name = "frmDrink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBegin_Click()
'once the user clicks the command button, the current form is hidden and the next form in the program appears
frmDrink.Hide
frmPick.Show
End Sub
