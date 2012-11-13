VERSION 5.00
Begin VB.Form frmFail 
   BackColor       =   &H000000FF&
   Caption         =   "Failure"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   3240
      Picture         =   "frmFail.frx":0000
      ScaleHeight     =   2895
      ScaleWidth      =   3135
      TabIndex        =   2
      Top             =   1080
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      Picture         =   "frmFail.frx":17DC
      ScaleHeight     =   2655
      ScaleWidth      =   2655
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Paul Bivens"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9240
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "We're sorry but your user name and/or password have been rejected. Please try again later."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmFail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sexton Dining Cash Register "\SextonDiningCashRegister.vpb"
'frmFail "/frmFail.frm"
'Paul Bivens
'March 22nd, 2006
'This form is used to prevent the user from using the program
'if the time runs out or they input 3 incorrect username or
'passwords.

'Ends the program
Private Sub cmdQuit_Click()
    End
End Sub
