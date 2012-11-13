VERSION 5.00
Begin VB.Form frmBartender 
   Caption         =   "The Bartender"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   Picture         =   "frmBartender.frx":0000
   ScaleHeight     =   6765
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Menu"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdShooters 
      Caption         =   "Shooters"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3240
      MaskColor       =   &H00000000&
      Picture         =   "frmBartender.frx":12ABF
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdFrozen 
      Height          =   1575
      Left            =   1680
      Picture         =   "frmBartender.frx":13610
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdMixed 
      BackColor       =   &H0000C000&
      Caption         =   "Mixed"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Picture         =   "frmBartender.frx":14952
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblAskMike 
      BackStyle       =   0  'Transparent
      Caption         =   "Ask the Bartender how to make your favorite Drinks!"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   4575
      Left            =   6360
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frmBartender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Bartending School
    'frmBartender(The Bartender)
    'By Fred Paul & Michael McKeever
    'March 22,2006
    'Welcome to the Bartender.  This form is used to explain to the
    'user how to make their favorite drinks.  Click on the buttons to
    'lead to separate forms.

Private Sub cmdBack_Click()
    'This button will hide the current form and Return the user to the
    'Bar form
        frmBartender.Hide
        frmBar.Show
End Sub

Private Sub cmdFrozen_Click()
    'Clicking this button will hide the bartender form and show
    'the frozen drinks form
    frmBartender.Hide
    frmFrozen.Show
End Sub

Private Sub cmdMixed_Click()
    'This Button will hide the bartender form and show the mixed
    'drinks form
    frmBartender.Hide
    frmMixed.Show
End Sub

Private Sub cmdShooters_Click()
    'This button will hide the bartender form and show the shooters
    'form
    frmBartender.Hide
    frmShooters.Show
End Sub

Private Sub Form_Load()
    'Since this was our original first form we needed to hide it
    'during the initial load process of the program, so it hides
    'the bartender form and shows the bar form.
    frmBartender.Hide
    frmBar.Show
End Sub


Private Sub lblAskMike_Click()

End Sub
