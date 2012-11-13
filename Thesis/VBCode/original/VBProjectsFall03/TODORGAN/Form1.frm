VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   675
   ClientTop       =   1065
   ClientWidth     =   4545
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   4545
   Begin VB.CommandButton Command4 
      Caption         =   "Quit"
      Height          =   975
      Left            =   840
      MaskColor       =   &H0000FFFF&
      TabIndex        =   5
      Top             =   5400
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Go Shopping!"
      Height          =   855
      Left            =   840
      MaskColor       =   &H0000FFFF&
      TabIndex        =   4
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find what SUV is right for you"
      Height          =   855
      Left            =   840
      MaskColor       =   &H0000FFFF&
      TabIndex        =   3
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Take a Look at the SUVs"
      Height          =   855
      Left            =   840
      MaskColor       =   &H0000FFFF&
      TabIndex        =   2
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   $"Form1.frx":0000
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Please select an option"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to TD's Luxury Autos"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Buying an SUV
'Form Name : Form1.frm
'Author: Tom Dorgan
'Date Written: October 28, 2003
'Purpose of Form: To direct the user to different forms on the site, each offering
                    'different options which help the user purchase a vehicle





Private Sub Command1_Click()
Form2.Hide
Form1.Hide
Form3.Show
Form4.Hide


End Sub

Private Sub Command2_Click()
Form3.Hide
Form1.Hide
Form2.Show
Form4.Hide

End Sub

Private Sub Command3_Click()
Form3.Hide
Form1.Hide
Form2.Hide
Form4.Show
End Sub

Private Sub Command4_Click()
End
End Sub
