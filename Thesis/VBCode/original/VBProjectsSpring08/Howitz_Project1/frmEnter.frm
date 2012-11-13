VERSION 5.00
Begin VB.Form frmEnter 
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuiit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3600
      TabIndex        =   5
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   4
      Top             =   5160
      Width           =   3015
   End
   Begin VB.TextBox txtpassword 
      Height          =   855
      Left            =   3720
      TabIndex        =   2
      Top             =   3960
      Width           =   3975
   End
   Begin VB.TextBox txtname 
      Height          =   855
      Left            =   3720
      TabIndex        =   1
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Label lblword 
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label lblname 
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   2325
      Left            =   2040
      Picture         =   "frmEnter.frx":0000
      Top             =   240
      Width           =   4500
   End
End
Attribute VB_Name = "frmEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sexton Cash Register
'Form Name:  frmEnter
'Louis Howitz
'March 31, 2008
'The purpose of this project is to create a functional cash register
'inspired by the till used at Sexton Dining at St. John's University.
'It includes a majority of the items available at Sexton.  This project
'is designed to list the items that the customer is purchasing and give
'them the total amount due.  This first form is designed as a login page
'that allows the user to operate the cash register.  I have only included
'two user names including mine and a "visitor" who can access the till.
'If the wrong username and/or password are entered, a message box
'will appear to tell the user to try again.

Private Sub cmdEnter_Click()

    Dim User As String
    Dim Code As Integer
    User = txtname.Text
    Code = txtpassword.Text
    
        If User = "Howitz" And Code = "56907" Then
            MsgBox "Success! Welcome to the exciting world of Sexton", , "Sexton Dining"
            frmEnter.Hide
            frmTill.Show
        ElseIf User = "Visitor" And Code = "12345" Then
            MsgBox "Success! Welcome to the exciting world of Sexton", , "Sexton Dining"
            frmEnter.Hide
            frmTill.Show
        Else
            MsgBox "Error, Please try again", , "Login"
        End If
   
    
        
End Sub

Private Sub cmdQuiit_Click()
    End
    
End Sub

