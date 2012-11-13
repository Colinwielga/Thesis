VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Games.Or.Cutt Game Rentals and Reviews"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
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
      Left            =   4080
      TabIndex        =   3
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
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
      Left            =   7680
      TabIndex        =   1
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdJoin 
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
      Left            =   480
      TabIndex        =   0
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reviews and News"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label lblMainTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Games.Or.Cutt:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   6750
      Left            =   0
      Picture         =   "frmRentalMain.frx":0000
      Top             =   0
      Width           =   10800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdJoin_Click()
    frmMain.Hide
    MsgBox "You must first register!", , "Please Register"
    frmRegister.Show
End Sub
Private Sub cmdLogin_Click()
Dim ChosenName As String
Dim ChosenPassword As String
    If ChosenName = UserName And ChosenPassword = Password Then
        UserName = InputBox("Please Enter User Name", "Enter User Name")
        Password = InputBox("Please Enter User Password", "Enter User Password")
        MsgBox "Thank You, You May Proceed", , "Congratualions"
        frmMain.Hide
        frmSelectWant.Show
    End If
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

