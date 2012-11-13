VERSION 5.00
Begin VB.Form frmHotel 
   Caption         =   "Hotel Welcome Menu"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   Picture         =   "frmHotel.frx":0000
   ScaleHeight     =   8880
   ScaleWidth      =   14790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000C0&
      Caption         =   "Quit"
      Height          =   615
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00800080&
      Caption         =   "Log-In"
      Height          =   615
      Left            =   9360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   6240
      MousePointer    =   3  'I-Beam
      TabIndex        =   2
      Top             =   6120
      Width           =   3015
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   6240
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Log-in to Continue"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   615
      Left            =   5520
      TabIndex        =   5
      Top             =   5160
      Width           =   4695
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Our Hotel"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   4080
      Width           =   5775
   End
   Begin VB.Label lblPassword 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password Here -->"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label lblUserName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Username Here -->"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   5640
      Top             =   5520
      Width           =   15
   End
End
Attribute VB_Name = "frmHotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Hotel Checkin
'Form: Hotel
'Authors: Ellen Jansen & Stuart Van Ess
'Date: March 28, 2008
'Purpose:   The purpose of this form is to verify that the user has proper
'           authority to work within the program
Option Explicit
Private Sub cmdEnter_Click()
'This shows the main menu of the program once you have
'entered the proper username and password, and then
'hides the log-in page.
    frmHotel.Hide
    frmMainMenu.Show
End Sub

Private Sub cmdOK_Click()


'This states that in order to log into the program, you
'must enter the proper Username and Password.
'If you do not, you are asked to try again.
    If (txtUserName.Text = "srvaness" And txtPassword.Text = "compsci") Or (txtUserName.Text = "ecjansen" And txtPassword.Text = "welcome") Then
        frmMainMenu.Show
        frmHotel.Hide
    Else
        MsgBox "Check For Incorrect Username or Password.", , "Sorry"
    End If
'This resets the username and password to be blank.
    txtUserName.Text = ""
    txtPassword.Text = ""
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

