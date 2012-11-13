VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00C000C0&
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   BeginProperty Font 
      Name            =   "Curlz MT"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "login.frx":0000
   ScaleHeight     =   7665
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnext 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "Click Here to Move on!"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      MaskColor       =   &H0080FFFF&
      Picture         =   "login.frx":3F9F0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Please Enter Your Full Name Below"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdnext_Click()
Bid.Show
Login.Hide
WholeName = txtname.Text
MsgBox "Welcome " & WholeName, , "HELLO"
End Sub

