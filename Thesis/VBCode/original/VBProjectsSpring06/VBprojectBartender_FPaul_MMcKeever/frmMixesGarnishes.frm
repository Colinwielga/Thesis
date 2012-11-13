VERSION 5.00
Begin VB.Form frmMixesGarnishes 
   Caption         =   "Mixes and Garnishes"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   Picture         =   "frmMixesGarnishes.frx":0000
   ScaleHeight     =   6300
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back "
      Height          =   615
      Left            =   6120
      TabIndex        =   8
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdGarnishes 
      Caption         =   "Garnishes"
      Height          =   615
      Left            =   6120
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCommonMixers 
      Caption         =   "Common Mixers"
      Height          =   615
      Left            =   6120
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrenadine 
      Caption         =   "Grenadine"
      Height          =   615
      Left            =   6120
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdBitters 
      Caption         =   "Bitters"
      Height          =   615
      Left            =   6120
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdSourMix 
      Caption         =   "Sour Mix"
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCream 
      Caption         =   "Cream"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdClubSoda 
      Caption         =   "Club Soda"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdTonicWater 
      Caption         =   "Tonic Water"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmMixesGarnishes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Bartending School
'frmMixesGarnishes(Mixes and Garnishes)
'By Fred Paul & Michael McKeever
'March 22,2006
'The Mixes Garnishes form defines common ingredients used in
'making the drinks specified in other forms.

Private Sub cmdBack_Click()
'This button when clicked returns the user to the bar tools page
'and hides the current form.
    frmMixesGarnishes.Hide
    frmBarTools.Show
End Sub
'all the following  buttons display what the button title claims
'in a msgbox.
Private Sub cmdBitters_Click()
    MsgBox "Bitters are secondary mixers used to settle the harsh taste of liquors.", , "Bitters"
End Sub

Private Sub cmdClubSoda_Click()
    MsgBox "Club Soda is carbonated water.", , "Club Soda"
End Sub

Private Sub cmdCommonMixers_Click()
    MsgBox "Other common mixers include: Coffee, Juices (Cranberry, Grapefruit, Lime/Lemon, Orange, and Pineapple), Sodas (7up/Sprite, Coke/Pepsi, and Ginger Ale).", , "Common Mixers"
End Sub

Private Sub cmdCream_Click()
    MsgBox "Cream otherwise called half and half, heavy cream, whole milk or low-fat milk.  Used to yield a hint of dairy flavor. ", , "Cream"
End Sub

Private Sub cmdGarnishes_Click()
    MsgBox "Garnishes are used to add flavor, and give eye appeal. Common garnishes include: Lemons, Limes, Celery, Coctail Onions, Maraschino Cherries, Olives, Oranges, Sugar, Tobasco Sauce, and Worcestershire Sauce.", , "Garnishes"
End Sub

Private Sub cmdGrenadine_Click()
    MsgBox "A sweet syrup mixer made from pomegranate juice.", , "Grenadine"
End Sub

Private Sub cmdSourMix_Click()
    MsgBox "A combination of lemon juice, lime juice and sugar.", , "Sour Mix"
End Sub

Private Sub cmdTonicWater_Click()
    MsgBox "Lemon and lime flavored quinine water.", , "Tonic Water"
End Sub
