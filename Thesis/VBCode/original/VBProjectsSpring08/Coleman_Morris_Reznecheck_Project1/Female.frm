VERSION 5.00
Begin VB.Form frm10 
   BackColor       =   &H0000FFFF&
   Caption         =   "Female"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10935
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdexecute 
      Caption         =   "Calculate "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtheight 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.PictureBox picoutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4440
      ScaleHeight     =   1155
      ScaleWidth      =   4155
      TabIndex        =   2
      Top             =   2400
      Width           =   4215
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to Ideal Weight"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Enter Your Height In Inches"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   5
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "frm10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Healthy Living
'frm10
'Ben Morris
'March 21
'calculate the ideal weight for a female

Private Sub cmdback_Click()
    'hides this form and opens the ideal weight form
    frm2.Show
    frm10.Hide
End Sub

Private Sub cmdclear_Click()
    picoutput.Cls
    'clears the output
End Sub

Private Sub cmdexecute_Click()

    Dim height As Single
    Dim idealweight As Single
    
    'this is the robinson formula for determaining ideal body weight
    height = txtheight.Text
    idealweight = 108.0265 + (height - 60) * 3.7479
    
    'prints the ideal weight as determained above
    picoutput.Print "Your Ideal Weight Is"
    picoutput.Print FormatNumber(idealweight), "lbs"
    
End Sub
