VERSION 5.00
Begin VB.Form BeginnerBinding 
   BackColor       =   &H000080FF&
   Caption         =   "Shop for bindings"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFischer 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   1440
      TabIndex        =   6
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdMarker 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   7080
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdRossi 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "Click on a binding for a description and a price!"
      Top             =   240
      Width           =   3495
   End
   Begin VB.PictureBox pbxMarker 
      Height          =   2295
      Left            =   6240
      Picture         =   "BeginnerBinding.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   240
      Width           =   4455
   End
   Begin VB.PictureBox pbxFischer 
      Height          =   3375
      Left            =   3000
      Picture         =   "BeginnerBinding.frx":5B8D
      ScaleHeight     =   3315
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   4560
      Width           =   4455
   End
   Begin VB.PictureBox pbxRossi 
      Height          =   2895
      Left            =   0
      Picture         =   "BeginnerBinding.frx":AAE7
      ScaleHeight     =   2835
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   840
      Width           =   3375
   End
End
Attribute VB_Name = "BeginnerBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: DavesSkiShop (Ski.vbp)
'Form name: BeginnerBinding (BeginnerBinding.frm)
'Author David Meacham
'Date Written: Wednesday, October 23
'Purpose of form: Allows user to find information about the binding displayed.
                    ' As well it allows them to purchase a pair of
                    'bindings, adding the cost to the running total and
                    'move on to the next form.

Option Explicit

Private Sub cmdFischer_Click()
'Adds price of binding to the running total and moves on to next form
sum = sum + 240
BeginnerBinding.Hide
BeginnerBoot.Show
End Sub

Private Sub cmdMarker_Click()
'Adds price of binding to the running total and moves on to next form
sum = sum + 110
BeginnerBinding.Hide
BeginnerBoot.Show
End Sub

Private Sub cmdRossi_Click()
'Adds price of binding to the running total and moves on to next form
sum = sum + 75
BeginnerBinding.Hide
BeginnerBoot.Show
End Sub

Private Sub pbxFischer_Click()
'Displays a description and price of the binding
MsgBox "This is Fischer's FX 10.  A great binding for a little more experienced beginner.  They are $240."
End Sub

Private Sub pbxMarker_Click()
'Displays a description and price of the binding
MsgBox "This is Marker's M1100 CCii.  A great beginner's binding.  They are $110."
End Sub

Private Sub pbxRossi_Click()
'Displays a description and price of the binding
MsgBox "This is Rossignol's Saphir binding.  Great for learners.  They are $75."
End Sub
