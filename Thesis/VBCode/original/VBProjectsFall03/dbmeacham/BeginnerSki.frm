VERSION 5.00
Begin VB.Form BeginnerSki 
   BackColor       =   &H00004080&
   Caption         =   "Shop for skiis"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Size 
      Caption         =   "Find your ski size here"
      Height          =   1095
      Left            =   4320
      TabIndex        =   7
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton cmdFischer 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   4440
      TabIndex        =   6
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdK2 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   8040
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdRossi 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C00000&
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Text            =   "Click on a ski to get a description and price!"
      Top             =   360
      Width           =   3255
   End
   Begin VB.PictureBox pbxFischer 
      Height          =   615
      Left            =   1560
      Picture         =   "BeginnerSki.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   7395
      TabIndex        =   2
      Top             =   6840
      Width           =   7455
   End
   Begin VB.PictureBox cmd 
      Height          =   7455
      Left            =   9600
      Picture         =   "BeginnerSki.frx":2466
      ScaleHeight     =   7395
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox pbxRossi 
      Height          =   7815
      Left            =   120
      Picture         =   "BeginnerSki.frx":5F59
      ScaleHeight     =   7755
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "BeginnerSki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: DavesSkiShop (Ski.vbp)
'Form name: BeginnerSki (BeginnerSki.frm)
'Author David Meacham
'Date Written: Wednesday, October 22
'Purpose of form: Allows user to find information about the ski's displayed.
                    'Allows them to search for their ski size.  As well it
                    'allows them to purchase a pair of skis and move on to
                    'the next form.
                    
Option Explicit

                    
Private Sub cmd_Click()
'Displays description and price of ski
MsgBox "This is k2's Axis SR.  A great beginner ski.  They are $375."
End Sub

Private Sub cmdFischer_Click()
'Adds the price of the ski to the running total and moves on to next form
sum = 0
sum = sum + 400
BeginnerSki.Hide
BeginnerBinding.Show
End Sub

Private Sub cmdK2_Click()
'Adds the price of the ski to the running total and moves on to next form
sum = 0
sum = sum + 375
BeginnerSki.Hide
BeginnerBinding.Show
End Sub

Private Sub cmdRossi_Click()
'Adds the price of the ski to the running total and moves on to next form
sum = 0
sum = sum + 350
BeginnerSki.Hide
BeginnerBinding.Show
End Sub



Private Sub pbxFischer_Click()
'Displays description and price of ski
MsgBox "This is Fischer's S200 RailFlex.  A great beginner ski.  They are $400."
End Sub

Private Sub pbxRossi_Click()
'Displays description and price of ski
MsgBox "This is Rossignol's Axium TD.  A great beginner ski.  They are $350."
End Sub

Private Sub Size_Click()
'This will allow the user to enter their height and find out what size ski they need
Dim Size As Integer
Size = InputBox("Enter your height in feet.  Please round to the nearest foot.")    'asks user to enter their height
Select Case Size
    Case Is <= 4                                                              'searches for the users height and prints out the corresponding ski size
        MsgBox "You need ski's between 160cm and 170cm."
    Case Is <= 5
        MsgBox "You need ski's between 180cm and 195cm."
    Case Is <= 6
        MsgBox "You need ski's between 200cm and 210cm."
    Case Else
        MsgBox "You need ski's 210cm or bigger."
End Select
        
End Sub
