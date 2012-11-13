VERSION 5.00
Begin VB.Form frm9 
   BackColor       =   &H0000FFFF&
   Caption         =   "Male"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
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
      Left            =   720
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.PictureBox picoutput 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3240
      ScaleHeight     =   1155
      ScaleWidth      =   4155
      TabIndex        =   4
      Top             =   2400
      Width           =   4215
   End
   Begin VB.TextBox txtheight 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
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
      Left            =   720
      TabIndex        =   1
      Top             =   960
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
      Left            =   720
      TabIndex        =   0
      Top             =   3840
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
      Height          =   735
      Left            =   3000
      TabIndex        =   3
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "frm9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Healthy Living
'frm9
'Ben Morris
'March 21
'calculates the ideal weight for a man
Private Sub cmdback_Click()
    frm2.Show
    frm9.Hide
End Sub

Private Sub cmdclear_Click()
    picoutput.Cls
End Sub

Private Sub cmdexecute_Click()
    Dim height As Single
    Dim idealweight As Single
    
    height = txtheight.Text
    idealweight = 114.64 + (height - 60) * 4.18888
    
    picoutput.Print "Your Ideal Weight Is"
    picoutput.Print FormatNumber(idealweight), "lbs"
    
End Sub
