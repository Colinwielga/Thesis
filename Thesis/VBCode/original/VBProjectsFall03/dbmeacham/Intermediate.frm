VERSION 5.00
Begin VB.Form IntermediateSki 
   Caption         =   "Form1"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFischer 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   4080
      TabIndex        =   6
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdK2 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   8160
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdRossi 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Text            =   "Click on a ski for a descpriton and a price!"
      Top             =   120
      Width           =   3135
   End
   Begin VB.PictureBox pbxFischer 
      Height          =   855
      Left            =   1320
      Picture         =   "Intermediate.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   7635
      TabIndex        =   2
      Top             =   6840
      Width           =   7695
   End
   Begin VB.PictureBox pbxK2 
      Height          =   7695
      Left            =   9360
      Picture         =   "Intermediate.frx":3822
      ScaleHeight     =   7635
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox pbxRossi 
      Height          =   7695
      Left            =   120
      Picture         =   "Intermediate.frx":712A
      ScaleHeight     =   7635
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "IntermediateSki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFischer_Click()
'Adds the price of the ski to the running total
sum = sum + 550
End Sub

Private Sub cmdK2_Click()
'Adds the price of the ski to the running total
sum = sum + 525
End Sub

Private Sub cmdRossi_Click()
'Adds the price of the ski to the running total
sum = sum + 500
End Sub

Private Sub pbxFischer_Click()
'Displays description and price of ski
MsgBox "This is Fischer's Airstyle NT.  An excellent freestyle ski.  They are $550."
End Sub

Private Sub pbxK2_Click()
'Displays description and price of ski
MsgBox "This is k2's 2500.  A great all mountain ski.  They are $525."
End Sub

Private Sub pbxRossi_Click()
'Displays description and price of ski
MsgBox "This is Rossignol's Bandit.  An Excellent all mountain ski.  They are $500."
End Sub
