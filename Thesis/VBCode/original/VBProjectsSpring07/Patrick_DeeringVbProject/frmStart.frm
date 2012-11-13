VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00000000&
   Caption         =   "Let's Go Shopping!"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H0000FFFF&
      Caption         =   "Start Shopping!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      MaskColor       =   &H0000FFFF&
      TabIndex        =   2
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   2655
   End
   Begin VB.PictureBox picGroceries 
      Height          =   1695
      Left            =   3720
      Picture         =   "frmStart.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      Caption         =   "By: Patrick Deering"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8040
      TabIndex        =   4
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00000000&
      Caption         =   "  Get Your Groceries        Delivered to You             on Campus!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2415
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'This program is set up to allow users to purchase any quantity of
'certain goods that he/she desires.  After recieving quantities,
'the program calculates the subtotal, delivery fee, tax and total
'that the user will be charged.  The programs ends by notifying
'the user of when the items will be delivered and where they
'will be delivered.

Private Sub CmdQuit_Click()
End
End Sub

Private Sub cmdStart_Click()
fullname = InputBox("Please input your name:", "Welcome!")
location = InputBox("Please input your dorm hall and number:", "Welcome!")
frmStart.Hide
frmHome.Show
'Retrieves user's name, then directs user to Home form
End Sub

