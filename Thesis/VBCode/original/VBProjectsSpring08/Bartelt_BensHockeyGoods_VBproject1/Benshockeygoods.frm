VERSION 5.00
Begin VB.Form frmDepartments 
   Caption         =   "Ben's Hockey Goods"
   ClientHeight    =   6960
   ClientLeft      =   9510
   ClientTop       =   4950
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "Goudy Old Style"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Benshockeygoods.frx":0000
   ScaleHeight     =   6960
   ScaleWidth      =   7650
   Begin VB.CommandButton cmdLeave 
      BackColor       =   &H000000FF&
      Caption         =   "Leave Ben's Hockey Goods"
      Height          =   855
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   2895
   End
   Begin VB.CommandButton cmdCheckOut 
      BackColor       =   &H000000FF&
      Caption         =   "Proceed to Checkout"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   2895
   End
   Begin VB.CommandButton cmdAccessories 
      BackColor       =   &H000000FF&
      Caption         =   "Accessories"
      Height          =   735
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdSkates 
      BackColor       =   &H000000FF&
      Caption         =   "Skates"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdHelmets 
      BackColor       =   &H000000FF&
      Caption         =   "Helmets"
      Height          =   735
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdSticks 
      BackColor       =   &H000000FF&
      Caption         =   "Sticks"
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdPadding 
      BackColor       =   &H000000FF&
      Caption         =   "Padding"
      Height          =   735
      Left            =   240
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "frmDepartments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ben's Hockey Store
'frmDepartments
'Ben Bartelt
'3/26/08
'This form shows all the the required equipment that is required to be a safe hockey player.
'It is a basically a home page for the user. Once they are done in one equipment section they
'then come back to this form. Once they have purchased equipment and returned to this form they are
'then allowed to go to the checkout form
'i have other comments under one of the buttons because its all the same idea.
Option Explicit
'Just a bunch of subroutines hiding and showing various departments.
Private Sub cmdCheckOut_Click()
'This button is not available until an item has been added to the cart of any department.
frmDepartments.Hide
frmCheckout.Show
End Sub

Private Sub cmdPadding_Click()
frmDepartments.Hide
frmPadding.Show
End Sub

Private Sub cmdAccessories_Click()
frmDepartments.Hide
frmAccessories.Show
End Sub

Private Sub cmdSticks_Click()
frmDepartments.Hide
frmSticks.Show
End Sub

Private Sub cmdHelmets_Click()
frmDepartments.Hide
frmHelmets.Show
End Sub

Private Sub cmdLeave_Click()
'This subroutine thanks the user for shopping via a msgbox, and quits the program.
MsgBox "Thanks for shopping at Ben's Ultimite Hockey Goods!  We hope you enjoyed our store.", , "Ben's Hockey Goods"
End
End Sub

Private Sub cmdSkates_Click()
frmDepartments.Hide
frmSkates.Show
End Sub



