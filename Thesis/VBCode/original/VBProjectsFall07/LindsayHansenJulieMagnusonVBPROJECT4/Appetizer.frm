VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H008080FF&
   Caption         =   "Form1"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTitlePicture 
      Height          =   4695
      Left            =   3720
      Picture         =   "Appetizer.frx":0000
      ScaleHeight     =   4635
      ScaleWidth      =   6195
      TabIndex        =   4
      Top             =   2160
      Width           =   6255
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quit!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton cmdDessert 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pick a Dessert!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdEntree 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pick an Entree!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdAppetizer 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pick an Appetizer!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Want some delicious recipes??? Then you've come to the right place!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1560
      TabIndex        =   5
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAppetizer_Click()
Appetizer.Show  'the purpose of Appetizer is to have the main page accesible throughout the whole program
Main.Hide       'we organized this so that one form would be up at a time
MsgBox "Pick Either an Asian Appetizer or an Italian Appetizer!"
End Sub


Private Sub cmdDessert_Click()
Dessert.Show    'this shows the dessert form
Main.Hide       'this hides the main form
MsgBox "Pick Either an Asian Dessert or an Italian Dessert!" 'this message box makes sure the user knows what to do on the form
End Sub

Private Sub cmdEntree_Click()
Entree.Show 'this shows the entree form
Main.Hide   'this hides the main form
MsgBox "Pick Either an Asian Entree or an Italian Entree!"  'this message box makes sure the user knows what to do on the form
End Sub

Private Sub cmdQuit_Click()
End
End Sub


