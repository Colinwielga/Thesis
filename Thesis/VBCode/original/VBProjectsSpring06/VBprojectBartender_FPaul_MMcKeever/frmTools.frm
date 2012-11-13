VERSION 5.00
Begin VB.Form frmBarTools 
   BackColor       =   &H80000007&
   Caption         =   "Tools of the Trade"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   Picture         =   "frmTools.frx":0000
   ScaleHeight     =   8460
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Menu"
      Height          =   615
      Left            =   8280
      TabIndex        =   2
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdGlasses 
      BackColor       =   &H8000000D&
      Caption         =   "Glasses"
      Height          =   615
      Left            =   8280
      TabIndex        =   1
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdMixed 
      Caption         =   "Mixes/Garnishes"
      Height          =   615
      Left            =   8280
      TabIndex        =   0
      Top             =   4560
      Width           =   1455
   End
End
Attribute VB_Name = "frmBarTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Bartending School
    'frmBarTools(Tools of the Trade)
    'By Fred Paul & Michael McKeever
    'March 22,2006
    'The Bar Tools form is used to teach the user about Mixes
    'and Garnishes, as well as different kinds of Glasses.
    'Click on their buttons to lead to their separate forms.

Private Sub cmdBack_Click()
    'This Button Hides the Bartools form and Reterns the user to
    'The Bar
    frmBarTools.Hide
    frmBar.Show
End Sub

Private Sub cmdGlasses_Click()
    'This Button hides the bartools form and Brings the user to the
    'Glasses form.
    frmBarTools.Hide
    frmGlasses.Show
End Sub

Private Sub cmdMixed_Click()
    'This Button Hides the Bartools form and redirects the user to
    'the mixesgarnishes form.
    frmBarTools.Hide
    frmMixesGarnishes.Show
End Sub


