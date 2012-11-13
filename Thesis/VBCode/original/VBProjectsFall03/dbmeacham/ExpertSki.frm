VERSION 5.00
Begin VB.Form ExpertSki 
   BackColor       =   &H00FF0000&
   Caption         =   "Shop for skiis"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Size 
      Caption         =   "Find your ski size here"
      Height          =   1095
      Left            =   4080
      TabIndex        =   7
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000000FF&
      Height          =   285
      Left            =   3600
      TabIndex        =   6
      Text            =   "Click on a ski to get a description and a price!"
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton cmdFischer 
      BackColor       =   &H00C00000&
      Caption         =   "Purchase"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   5880
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   1080
      Picture         =   "ExpertSki.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   8235
      TabIndex        =   4
      Top             =   6360
      Width           =   8295
   End
   Begin VB.CommandButton cmdk2 
      Caption         =   "Purchase"
      Height          =   495
      Left            =   8400
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cdmRossi 
      Caption         =   "Purchase"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox k2 
      Height          =   7455
      Left            =   9480
      Picture         =   "ExpertSki.frx":3414
      ScaleHeight     =   7395
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.PictureBox Scratch 
      Height          =   7215
      Left            =   0
      Picture         =   "ExpertSki.frx":6F15
      ScaleHeight     =   7155
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   -120
      Width           =   975
   End
End
Attribute VB_Name = "ExpertSki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: DavesSkiShop (Ski.vbp)
'Form name: ExpertSki (ExpertSki.frm)
'Author David Meacham
'Date Written: Wednesday, October 22
'Purpose of form: Allows user to find information about the ski's displayed.
                    'Allows them to search for their ski size.  As well it
                    'allows them to purchase a pair of skis and move on to
                    'the next form.
    
Option Explicit

Private Sub cdmrossi_Click()
'Adds the cost of the ski to the running total and moves on to next form
sum = 0
sum = sum + 800
ExpertSki.Hide
ExpertBinding.Show
End Sub

Private Sub cmdFischer_Click()
'Adds the cost of the ski to the running total and moves on to next form
sum = 0
sum = sum + 795
ExpertSki.Hide
ExpertBinding.Show
End Sub

Private Sub cmdK2_Click()
'Adds the cost of the ski to the running total and moves on to next form
sum = 0
sum = sum + 750
ExpertSki.Hide
ExpertBinding.Show
End Sub




Private Sub k2_Click()
'Displays a description and price of the ski.
MsgBox "These are the k2 5500.  An excellent all mountain ski.  They are $750."
End Sub

Private Sub Picture1_Click()
'Displays a description and price of the ski.
MsgBox "This is Fischer's Big Stix FX 10.6.  An excellent racing ski.  They are $795."
End Sub

Private Sub Scratch_Click()
'Displays a description and price of the ski.
MsgBox "These are Rossignol's Scratch.  An excellent freestyle ski.  They are $800."
End Sub

Private Sub Size_Click()
'This will allow the user to enter their height and find out what size ski they need
Dim Size As Integer
Size = InputBox("Enter your height in feet.  Please round to the nearest foot.")        'asks the user to enter their height
If Size <= 4 Then
        MsgBox "You need ski's between 160cm and 170cm."                'searches for the users height and prints out the corresponding ski size
    ElseIf Size <= 5 Then
        MsgBox "You need ski's between 180cm and 195cm."
    ElseIf Size <= 6 Then
        MsgBox "You need ski's between 200cm and 210cm."
    Else
        MsgBox "You need ski's 210cm or bigger."
End If
End Sub

