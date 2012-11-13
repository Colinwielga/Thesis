VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FF0000&
   Caption         =   "Form2"
   ClientHeight    =   8655
   ClientLeft      =   2355
   ClientTop       =   1125
   ClientWidth     =   9885
   LinkTopic       =   "Form2"
   ScaleHeight     =   8655
   ScaleWidth      =   9885
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Choose one of the options then Click me "
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      TabIndex        =   8
      Top             =   6720
      Width           =   2655
   End
   Begin VB.OptionButton optmore 
      Caption         =   "More than Five"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6240
      TabIndex        =   6
      Top             =   5040
      Width           =   2055
   End
   Begin VB.OptionButton optfive 
      Caption         =   "       Five"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6240
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.OptionButton optfour 
      Caption         =   "      Four"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6240
      TabIndex        =   4
      Top             =   2280
      Width           =   2055
   End
   Begin VB.OptionButton optthree 
      Caption         =   "     Three"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   3
      Top             =   5040
      Width           =   2055
   End
   Begin VB.OptionButton opttwo 
      Caption         =   "       Two"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   2
      Top             =   3600
      Width           =   2055
   End
   Begin VB.OptionButton optone 
      Caption         =   "       One"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Choose one of the following:"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   3360
      TabIndex        =   7
      Top             =   1320
      Width           =   3555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "How much alcohol have you consumed ?"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   2505
      TabIndex        =   0
      Top             =   600
      Width           =   5490
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Blood Alcohol Content (Bloodalcohollevel.vbp)
'Form Name : Blood Alcohol Content (Number.frm)
'Author : Jason Sangworn
'Date Written : March 15, 2004
'Purpose of Project : To calculate the blood alcohol content
                      'from the user by using the
                      'type of alcohol consumed
                      'number of drinks consumed
                      'time in which the drinks were consumed
                      'weight of the user
                      'and finally calculate the BAC of the user
                      
'Purpose of Form : 'have the user click on the number of drinks consumed
                   'and arrange that information into arrays
                   'then proceed to the next form
Option Explicit


Private Sub Form_Load()
ReDim Number(1 To 6) As Double


'Open the data file "Number.txt " for the Arrays that
Open "N:\CS130\handin\Sangworn, Jason\Number.txt" For Input As #1
For J = 1 To 6
    Input #1, Number(J) 'Read data into the respective Arrays.
Next J
Close #1

End Sub


Private Sub Command1_Click()
If optone = True Then
    J = 1
ElseIf opttwo = True Then
    J = 2
ElseIf optthree = True Then
    J = 3
ElseIf optfour = True Then
    J = 4
ElseIf optfive = True Then
    J = 5
ElseIf optmore = True Then
    J = 6
End If

Form2.Hide
Form3.Show

End Sub
