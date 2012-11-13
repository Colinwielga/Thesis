VERSION 5.00
Begin VB.Form frmQ2 
   BackColor       =   &H00008000&
   Caption         =   "Question 2"
   ClientHeight    =   8505
   ClientLeft      =   1695
   ClientTop       =   1095
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   12345
   Begin VB.CommandButton cmdQ2B 
      BackColor       =   &H00008000&
      Caption         =   "Go Back To Main Menu"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdQ2A 
      BackColor       =   &H00008000&
      Caption         =   "Next Question"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdQ2C 
      BackColor       =   &H00008000&
      Caption         =   "Submit Answer"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   1395
   End
   Begin VB.TextBox txtA2 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9120
      TabIndex        =   2
      Top             =   2520
      Width           =   615
   End
   Begin VB.PictureBox pic3 
      Height          =   4815
      Left            =   2040
      Picture         =   "frmQ2.frx":0000
      ScaleHeight     =   4755
      ScaleWidth      =   3795
      TabIndex        =   1
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label lblMe 
      BackColor       =   &H00008000&
      Caption         =   "Lance Uselman"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label lblInput 
      BackColor       =   &H00008000&
      Caption         =   "Input correct number of years"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblQu2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   $"frmQ2.frx":463A
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   7455
   End
End
Attribute VB_Name = "frmQ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Wildlife Challenge (Project1.vbp)
'frmQ2 (frmQ2.frm)
'Lance Uselman
'March 24, 2006
'Purpose of the form: This form asks the user a question and to input the
'                     answer in a textbox. The appropriate message box follows.

Option Explicit
Private Sub cmdQ2A_Click()
    frmQ3.Show
    frmQ2.Hide  'This button allows the user to go to the next question.
End Sub

Private Sub cmdQ2B_Click()
    frmMain.Show
    frmQ2.Hide  'This button allows the user to go to the main form.
End Sub

Private Sub cmdQ2C_Click()
    Dim Age As Integer  'This button allows the user to submit the answer and then displays the appropriate message box.
    Age = txtA2.Text
    If Age <> 5 Then
        MsgBox "This age is incorrect. Hint: The correct number of years is somewhere between 3 and 7.", , "Wrong Answer"
    Else
        MsgBox "This is the correct answer.", , "Correct!"
    End If
End Sub
