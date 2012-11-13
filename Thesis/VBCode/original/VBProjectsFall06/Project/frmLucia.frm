VERSION 5.00
Begin VB.Form frmLucia 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   5520
      Width           =   3855
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit this form"
      Height          =   615
      Left            =   3000
      TabIndex        =   9
      Top             =   7320
      Width           =   2055
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   615
      Left            =   720
      TabIndex        =   8
      Top             =   7320
      Width           =   2055
   End
   Begin VB.TextBox txtQuestion 
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   7
      Top             =   6720
      Width           =   7695
   End
   Begin VB.TextBox txtCityState 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   6120
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   2040
      Picture         =   "frmLucia.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H0000FFFF&
      Caption         =   "Question for Coach Lucia"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   6
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label lblCity 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "City, State"
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label lblName 
      BackColor       =   &H0000FFFF&
      Caption         =   "Name"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   $"frmLucia.frx":40DF
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   480
      TabIndex        =   2
      Top             =   4680
      Width           =   8805
   End
   Begin VB.Label lblDon 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "The Don Lucia Show"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3405
      TabIndex        =   1
      Top             =   4200
      Width           =   2685
   End
End
Attribute VB_Name = "frmLucia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmLucia
'Cole and John
'10/30/06
'Objective: The objective of this form is to allow the user to ask head coach Don
'Lucia a question that he will answer on his weekly talk show.  The user inputs his/
'her name, hometown, and question into a textbox.  This information is then
'"submitted" to Coach Lucia for consideration.

Option Explicit

Private Sub cmdHome_Click()
    frmLucia.Visible = False
    frmMain.Visible = True
End Sub

Private Sub cmdSubmit_Click()
Dim Name As String          'declares Name variable as a word

    Name = txtName.Text     'Input in the text box is the Name variable

    MsgBox "Thank you " & Name & " for submitting the question for Coach Lucia", , "Question Submitted"
End Sub


