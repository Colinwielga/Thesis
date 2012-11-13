VERSION 5.00
Begin VB.Form frmFixings 
   BackColor       =   &H00FF00FF&
   Caption         =   "Fixings"
   ClientHeight    =   7140
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRice 
      Caption         =   "RICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6120
      Picture         =   "frmFixings.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdBBeans 
      Caption         =   "BLACK BEANS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6120
      Picture         =   "frmFixings.frx":0567
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdOrder2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Go Back to Order Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdSourCream 
      Caption         =   "SOUR CREAM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4320
      Picture         =   "frmFixings.frx":2831
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdLettuce 
      Caption         =   "LETTUCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2520
      Picture         =   "frmFixings.frx":700D
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdPBeans 
      Caption         =   "PINTO BEANS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      Picture         =   "frmFixings.frx":B379
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdCheese 
      Caption         =   "CHEESE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4320
      Picture         =   "frmFixings.frx":10E86
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalsa 
      Caption         =   "SALSA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2520
      Picture         =   "frmFixings.frx":11EBC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuac 
      Caption         =   "GUACAMOLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      Picture         =   "frmFixings.frx":2A020
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Go Back to Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FF00FF&
      Caption         =   "Designed by Carrie Hyland"
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label lblFixings 
      BackColor       =   &H00FF00FF&
      Caption         =   "Click on each picture to get a description of the fixings that come on any Chipotle product!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   9
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmFixings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: ProjChipotleOrder (Carrie Hyland's VB Project.vbp)
'Form Name: frmFixings (frmFixings_form.frm)
'Author: Carrie Hyland
'Date Written: October 19, 2003
'Purpose of Form: To allow the user to view the different
                 ' ingredients that can be put on either
                 ' a burrito, taco or bol.
                 
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit

Private Sub cmdBBeans_Click()
'A message pops up to inform the user about Black Beans.
MsgBox "Our black beans are never refried, but spiced up with cumin, garlic and other spices.  Always tender, never mushy!"
End Sub
Private Sub cmdCheese_Click()
'A message pops up to inform the user about cheese.
MsgBox "A unique blend of jack and chedder cheese.  The cheese is grated daily to ensure fresh, moist cheese."

End Sub

Private Sub cmdLettuce_Click()
'A message pops up to inform the user about lettuce.
MsgBox "Crisp, fresh romaine lettuce leaves add a clean and organic touch to any burrito!"

End Sub

Private Sub cmdMainMenu_Click()
'Hides the frmFixings and shows the frmMainMenu
'(switches from the fixings form to the Main Menu form).
frmFixings.Hide
frmMainMenu.Show

End Sub

Private Sub cmdOrder2_Click()
'Hides the frmFixings and shows the frmOrder
'(switches from the fixings form to the Main Menu form).
cmdOrder2.Enabled = True
frmFixings.Hide
frmOrder.Show

End Sub

Private Sub cmdPBeans_Click()
'A message pops up to inform the user of Pinto Beans.
MsgBox "The pinto beans have a smokey flavor with a pinch of spice.  They are also slow cooked with bacon."

End Sub

Private Sub cmdQuac_Click()
'A message pops up to inform the user about Guacamole.
MsgBox "Several times a day, Chipotle mashes fresh, ripe avocados and mixes them with cilantro, jalepeno peppers, lime juice and spices.  Guacamole is $1.25 extra."

End Sub

Private Sub cmdQuit_Click()
'Ends the program
End
End Sub

Private Sub cmdRice_Click()
'A message pops up to inform the user about rice.
MsgBox "Steamed fresh every hour and seasoned with fresh cilantro, dash of salt and a splash of fresh citrus juice."
End Sub

Private Sub cmdSalsa_Click()
'A message pops up to inform the user about Salsa.
MsgBox "Chopped, ripe, red tomatoes are combined with red onions, jalapeno peppers and cilantro for a refreshing taste."

End Sub

Private Sub cmdSourCream_Click()
'A message pops up to inform the user about sour cream.
MsgBox "One hundred percent fresh cultured cream whipped to a smooth consistency."

End Sub

