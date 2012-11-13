VERSION 5.00
Begin VB.Form frmFindBudget 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Best Fit Program by Budget"
   ClientHeight    =   6900
   ClientLeft      =   2940
   ClientTop       =   2025
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   10260
   Begin VB.PictureBox picMoneyGuy 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   6960
      Picture         =   "frmFindBudget.frx":0000
      ScaleHeight     =   3375
      ScaleWidth      =   2535
      TabIndex        =   5
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Click Here to Find Your Program"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox txtEnter 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblEnter 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enter Your Estimated Spending Allowance:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   2
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Best Fit Program by Budget"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10335
   End
End
Attribute VB_Name = "frmFindBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'written by Erika 3/24/08
' This subroutine asks the user for their expected spending allowance, then displays which programs they can afford
'in a message box
Private Sub cmdCalculate_Click()

Dim Budget As Single

Budget = txtEnter.Text

Select Case Budget
    Case Is < 2500
        MsgBox "Sorry, you can't afford to go abroad at this time, maybe you should look into going next year!", , "Sorry!"
    Case Is < 3000
        MsgBox "You can afford to go to Ireland at this time!", , "Ireland"
    Case Is < 3500
        MsgBox "You can have your pick of Ireland, France, Austria, or Spain! Good luck deciding!", , "Ireland, France, Austria, Spain"
    Case 4000 To 4999
        MsgBox "You can have your pick of Ireland, France, Austria, Spain, or the Greco-Roman trip! All are fun!", , "WHOA! Lots of choices!"
    Case Is >= 5000
        MsgBox "You get your pick of any program including the rather pricy London!", , "You can go anywhere!"
    Case Else
        MsgBox "Please enter a different budget.", , Error
End Select
 
End Sub

Private Sub cmdGoBack_Click()
frmFindBudget.Hide
frmFind.Show
End Sub


