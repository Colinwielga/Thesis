VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form2"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   LinkTopic       =   "Form2"
   ScaleHeight     =   2895
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3000
      ScaleHeight     =   855
      ScaleWidth      =   3135
      TabIndex        =   4
      Top             =   1200
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Back"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdID 
      Caption         =   "Create ID Tag"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your First Name, Middle Initial and Last Name:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Form2.Hide
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdID_Click()
Dim Name As String
Dim N As Integer
Dim First As String, Middle As String, Last As String

Name = txtName.Text         'Creates an ID tag with the first 3 letters of the first name, the middle initial, and the whole last name
N = InStr(Name, " ")        'The ID is a global variable so it can be used on the final ticket.
First = Left(Name, N - 1)
Last = Right(Name, Len(Name) - (N + 2))
Middle = Mid(Name, N + 1, 1)
ID = Left(First, 3) & Middle & Left(Last, 10)
picResults.Print " Your Name Tag is "; ID

End Sub

