VERSION 5.00
Begin VB.Form frmMessage 
   Caption         =   "Crafting Diplomacy in Writ"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   Picture         =   "frmMessage.frx":0000
   ScaleHeight     =   8625
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMsg4 
      Caption         =   "4. You may have my sister's hand if you cease your actions."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      TabIndex        =   5
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton cmdMsg2 
      Caption         =   "2.What is your price?"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton cmdMsg1 
      Caption         =   "1.This is WAR!"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CommandButton cmdMsg3 
      Caption         =   "3. My Lord, have I offended thee?"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      TabIndex        =   2
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
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
      Height          =   615
      Left            =   7320
      TabIndex        =   1
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   7215
      Left            =   0
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H0080FFFF&
      Caption         =   "It seems to me that you have one of four options my lord.  Choose wisely..."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'objective: to give the user options through command buttons given the situation described
'in the label
'certain decisions affect the battlepoints variable
'Each decision choice corresponds to the appearance of one of two forms.


Private Sub cmdMsg1_Click()
Battlepoints = Battlepoints + 200
frmMessage.Hide
frmMsgresponse.Show
End Sub

Private Sub cmdMsg2_Click()
Battlepoints = Battlepoints - 100
frmMessage.Hide
frmWar.Show
End Sub

Private Sub cmdMsg3_Click()
Battlepoints = Battlepoints + 100
frmMessage.Hide
frmWar.Show
End Sub

Private Sub cmdMsg4_Click()
Battlepoints = Battlepoints - 200
frmMessage.Hide
frmWar.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub
