VERSION 5.00
Begin VB.Form ExpertBinding 
   BackColor       =   &H0000FFFF&
   Caption         =   "Shop for bindings"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF00FF&
      Height          =   285
      Left            =   5880
      TabIndex        =   6
      Text            =   "Click on a binding to get a description and a price!"
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton cmdRossi 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdFischer 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   8040
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdMarker 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   3120
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.PictureBox Picture3 
      Height          =   2535
      Left            =   240
      Picture         =   "ExpertBinding.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   960
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   5040
      Picture         =   "ExpertBinding.frx":5EAE
      ScaleHeight     =   3555
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   2160
      Width           =   4455
   End
   Begin VB.PictureBox Picture2 
      Height          =   2895
      Left            =   240
      Picture         =   "ExpertBinding.frx":AD9C
      ScaleHeight     =   2835
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   4320
      Width           =   3375
   End
End
Attribute VB_Name = "ExpertBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: DavesSkiShop (Ski.vbp)
'Form name: ExpertBinding (ExpertBinding.frm)
'Author David Meacham
'Date Written: Wednesday, October 23
'Purpose of form: Allows user to find information about the binding displayed.
                    ' As well it allows them to purchase a pair of
                    'bindings, adding the cost to the running total and
                    'move on to the next form.

Option Explicit

Private Sub cmdFischer_Click()
'Adds cost of bindings to the running sum and moves on to next form
sum = sum + 310
ExpertBinding.Hide
ExpertBoot.Show
End Sub

Private Sub cmdMarker_Click()
'Adds cost of bindings to the running sum and moves on to next form
sum = sum + 350
ExpertBinding.Hide
ExpertBoot.Show
End Sub

Private Sub cmdRossi_Click()
'Adds cost of bindings to the running sum and moves on to next form
sum = sum + 350
ExpertBinding.Hide
ExpertBoot.Show
End Sub

Private Sub Picture1_Click()
' Displays description and price of binding.
MsgBox "These are Fischer's FR 17.  They are an excellent racing binding. They are $310."
End Sub

Private Sub Picture2_Click()
'Displays description and price of binding.
MsgBox "These are Rossignol's Scratch series.  They are excellent for freestyle skiing.  They are $300."
End Sub

Private Sub Picture3_Click()
' Displays description and price of binding.
MsgBox "These are the Marker Comp 1400.  They are excellent for all around skiing.  They are $350."
End Sub
