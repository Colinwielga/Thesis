VERSION 5.00
Begin VB.Form frmEntrance 
   Caption         =   "Ben's Hockey Goods"
   ClientHeight    =   7755
   ClientLeft      =   7035
   ClientTop       =   4725
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   Picture         =   "frmEntrance.frx":0000
   ScaleHeight     =   7755
   ScaleWidth      =   12405
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Leave"
      DisabledPicture =   "frmEntrance.frx":202CD
      DownPicture     =   "frmEntrance.frx":2229B
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8400
      MaskColor       =   &H8000000F&
      Picture         =   "frmEntrance.frx":24077
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton cmdBeginShopping 
      Caption         =   "Begin Shopping"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1680
      MaskColor       =   &H8000000F&
      Picture         =   "frmEntrance.frx":2D05F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Image imagepuck 
      Height          =   1215
      Left            =   5280
      Picture         =   "frmEntrance.frx":351B7
      Top             =   5160
      Width           =   1980
   End
   Begin VB.Label LblStore 
      BackColor       =   &H00C0FFFF&
      Caption         =   "How expensive is Hockey equipment?"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   7455
   End
   Begin VB.Shape shpStore 
      FillColor       =   &H0000FFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1335
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   8055
   End
End
Attribute VB_Name = "frmEntrance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ben's Hockey Store
'frmEntrance
'Ben Bartelt
'3/26/08
'The overall purpose of this project is to show how expensive it cost just to get a person fully equiped to play hockey.
'Just imagine if you start your kid young when he is growing. A person has to repurchase equipment every year because
'their prevous equipment is too small.
'The purpose of this form is just and entrance form like a store from. Allowing the shopper to leave or enter the store
'other comments are listed under subroutines
Option Explicit

Private Sub cmdBeginShopping_Click()
Dim X As Integer
frmEntrance.Hide
frmDepartments.Show
'it also opens each department's cart and erases the file by printing blanks in locations 0 to 100.
'This reenables the display button in the checkout form if this is a second shopping trip.
frmCheckout.cmdDisplay.Enabled = True

Open App.Path & "\SkatesCart.txt" For Output As #1

For X = 1 To 200
    Print #1, ""
Next X
Close #1

Open App.Path & "\HelmetsCart.txt" For Output As #2

For X = 1 To 200
    Print #2, ""
Next X
Close #2

Open App.Path & "\SticksCart.txt" For Output As #3

For X = 1 To 200
    Print #3, ""
Next X
Close #3

Open App.Path & "\PaddingCart.txt" For Output As #4

For X = 1 To 200
    Print #4, ""
Next X
Close #4

Open App.Path & "\AccessoriesCart.txt" For Output As #5

For X = 1 To 200
    Print #5, ""
Next X
Close #5


End Sub

Private Sub cmdQuit_Click()
'This subroutine quits out of the program and also sends the user a msgbox.
MsgBox "Have a wonderful day!", , "Ben's Hockey Goods"
End

End Sub

Private Sub Form_Load()
'This subroutine greats the user with a msgbox and the Entrance form.
frmEntrance.Show
MsgBox "Hello! Welcome to Ben's Hockey Goods, thank you for choosing us for your hockey needs. An item from each department display must be purchased to have proper protection.", , "Ben's Hockey Goods"

End Sub

Private Sub imagepuck_Click()
'image of a hockey puck
End Sub
