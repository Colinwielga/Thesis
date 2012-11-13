VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000003&
   Caption         =   "Form2"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6225
   LinkTopic       =   "Form2"
   ScaleHeight     =   2640
   ScaleWidth      =   6225
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm"
      Height          =   495
      Left            =   4800
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtInputD 
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtInputCD 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtInputC 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtInputBC 
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtInputB 
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtInputAB 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtInputA 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Yumin Lu - CS130"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Grade D"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Grade CD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Grade C"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Grade BC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Grade B"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Grade AB"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Grade A "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a new grading Scale - Number from 0 to 300"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Grading Program
'Yumin Lu
'March 4th
'Purpose:
'This program will conclude 3 test scores of students,
'including 2 tests and 1 final. It will ask the user for a desired grading scale and print out the grade of each
'student accordingly. The program will also include sorting and searching function.
'The grading method applies the rule that it automatically withdraw the lowest score that is other than the
'final while the final will be double-counted. As well, if the final is the lowest, the total will simply be the
'sum of all three scores.


Private Sub cmdConfirm_Click()
'Let user to fill up 7 numbers for the bottom line score of grade A, AB, B, BC, C, CD, D, the number shall be in the range from 0 to 300
'Remember the scale user has set up


A = txtInputA.Text
AB = txtInputAB.Text
B = txtInputB.Text
BC = txtInputBC.Text
C = txtInputC.Text
CD = txtInputCD.Text
D = txtInputD.Text

Form2.Hide

Form1.Show

'Make the GRADE button available
Form1.cmdGrade.Enabled = True


End Sub

