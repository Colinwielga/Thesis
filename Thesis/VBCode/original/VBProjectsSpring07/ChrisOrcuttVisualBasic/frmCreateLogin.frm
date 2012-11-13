VERSION 5.00
Begin VB.Form frmRegister 
   Caption         =   "New User"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
   End
   Begin VB.CommandButton cmdReturnMain 
      Caption         =   "Main Page"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblDirections 
      BackStyle       =   0  'Transparent
      Caption         =   "New user? Please register a new account!"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   7815
   End
   Begin VB.Label lblEnterName 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a user name:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label lblEnterPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a password:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   11220
      Left            =   0
      Picture         =   "frmCreateLogin.frx":0000
      Top             =   0
      Width           =   8790
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UserName As String, Password As String
Private Sub cmdRegister_Click()
Dim ChosenName As String
Dim ChosenPassword As String
    UserName = txtUserName.Text
    Password = txtPassword.Text
        MsgBox ("User name is: " & UserName & " Password is: " & Password)
    frmRegister.Hide
    frmMain.Show
End Sub
Private Sub cmdReturnMain_Click()
    frmRegister.Hide
    frmMain.Show
End Sub

