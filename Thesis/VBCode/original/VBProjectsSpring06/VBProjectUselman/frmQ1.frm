VERSION 5.00
Begin VB.Form frmQ1 
   BackColor       =   &H00008000&
   Caption         =   "Question 1"
   ClientHeight    =   8565
   ClientLeft      =   1365
   ClientTop       =   1095
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   12435
   Begin VB.CommandButton cmdQ1B 
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
      Index           =   1
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdQ1C 
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
      Index           =   0
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdQ1A 
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
      Height          =   615
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtA1 
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
      Height          =   375
      Left            =   10320
      TabIndex        =   6
      Top             =   3240
      Width           =   495
   End
   Begin VB.PictureBox picDuck 
      Height          =   3135
      Left            =   3960
      Picture         =   "frmQ1.frx":0000
      ScaleHeight     =   1777.457
      ScaleMode       =   0  'User
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   2160
      Width           =   5055
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
      TabIndex        =   11
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label lblInput 
      BackColor       =   &H00008000&
      Caption         =   "Input Number Of Correct Answer"
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
      Index           =   4
      Left            =   9840
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblD3 
      BackColor       =   &H00008000&
      Caption         =   "3: Lesser Scaup"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   600
      TabIndex        =   5
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label lblD4 
      BackColor       =   &H00008000&
      Caption         =   "4: Blue-Winged Teal"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label lblD2 
      BackColor       =   &H00008000&
      Caption         =   "2: Common Merganser"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label lblD1 
      BackColor       =   &H00008000&
      Caption         =   "1: Northern Pintail"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label lblQu1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "What is the name of this duck?"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   840
      Width           =   6975
   End
End
Attribute VB_Name = "frmQ1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Wildlife Challenge (Project1.vbp)
'frmQ1 (frmQ1.frm)
'Lance Uselman
'March 24, 2006
'Purpose of the form: This form asks the user to answer a question by typing the
'                     number associated with the answer into a textbox and then
'                     clicking on the button which submits the answer. The appropriate
'                     message box follows.

Option Explicit
Private Sub cmdQ1A_Click()
    Dim Answer As Single
    Answer = txtA1.Text 'This step gets input from textbox.
    Select Case Answer  'This step reads inputted number and sorts through case conditions to find associated value.
        Case Is = 1
            MsgBox "Northern Pintail is not the correct answer.", , "Wrong Answer"
        Case Is = 2
            MsgBox "Common Merganser is not the correct answer.", , "Wrong Answer"
        Case Is = 3
            MsgBox "Lesser Scaup is not the correct answer.", , "Wrong Answer"
        Case Is = 4
            MsgBox "Blue-Winged Teal is the correct answer. A unique feature of the Blue-Winged Teal is the white crescent-shaped patch on the side of the males' head.", , "Correct!"
    End Select
End Sub
Private Sub cmdQ1B_Click(Index As Integer)
    frmQ2.Show
    frmQ1.Hide  'This button allows the user to go to the next question.
End Sub
Private Sub cmdQ1C_Click(Index As Integer)
    frmMain.Show
    frmQ1.Hide  'This button allows the user to go back to the main form.
End Sub

