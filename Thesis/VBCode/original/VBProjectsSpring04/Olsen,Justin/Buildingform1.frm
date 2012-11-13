VERSION 5.00
Begin VB.Form Buildingform1 
   BackColor       =   &H00FF0000&
   Caption         =   "Form2"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form2"
   ScaleHeight     =   5280
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdlink1 
      BackColor       =   &H0000FF00&
      Caption         =   "Go to My Links!"
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdbegin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back to the Beginning."
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   120
      Picture         =   "Buildingform1.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   3360
      Picture         =   "Buildingform1.frx":BE27
      ScaleHeight     =   3315
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdbuild 
      BackColor       =   &H000000C0&
      Caption         =   "If you want to see the real deal about building your own ""Stripper"" CLICK HERE!"
      Height          =   855
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   4095
   End
   Begin VB.CommandButton cmdbuy 
      BackColor       =   &H0000C000&
      Caption         =   "If you have money to spend and no time to build your own cedar strip canoe, click on ME!!!"
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Designed by: Justin Olsen"
      Height          =   255
      Left            =   8520
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Buildingform1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Purpose = This form is here to ask the user if they would like to move on to a form about building canoes or buying a canoe.
'Project Name= Visual Basic Canoe Project("M:\CS130\CanoeProject")
'Form Name= Form 1 ("M:\CS130\CanoeProject\Project1.vbp\Buildingform1.frm")
Private Sub cmdbegin_Click()
    Form1.Show
    Buildingform1.Hide
End Sub

Private Sub cmdbuild_Click()
Buildingform2.Show
Buildingform1.Hide
End Sub

Private Sub cmdbuy_Click()
Buyingform.Show
Buildingform1.Hide
End Sub

Private Sub cmdlink1_Click()
Form2.Show
Buildingform1.Hide
End Sub

Private Sub cmdquit_Click()
End
End Sub
