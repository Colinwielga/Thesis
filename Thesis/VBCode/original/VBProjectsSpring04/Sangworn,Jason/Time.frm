VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FF0000&
   Caption         =   "Form3"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   1200
   ClientWidth     =   9885
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form3"
   ScaleHeight     =   9435
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
      Height          =   855
      Left            =   3480
      MaskColor       =   &H8000000F&
      TabIndex        =   6
      Top             =   7080
      Width           =   2775
   End
   Begin VB.OptionButton Option5 
      Caption         =   " More than 5"
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
      Left            =   1920
      TabIndex        =   5
      Top             =   5160
      Width           =   1935
   End
   Begin VB.OptionButton Option4 
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
      Left            =   4920
      TabIndex        =   4
      Top             =   3720
      Width           =   1935
   End
   Begin VB.OptionButton Option3 
      Caption         =   "      Three"
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
      Left            =   1920
      TabIndex        =   3
      Top             =   3720
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      Caption         =   "      Two"
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
      Left            =   4920
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "      One"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
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
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   3555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "How many hours have you been drinking ?"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   5745
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Blood Alcohol Content (Bloodalcohollevel.vbp)
'Form Name : Blood Alcohol Content (Time.frm)
'Author : Jason Sangworn
'Date Written : March 15, 2004
'Purpose of Project : To calculate the blood alcohol content
                      'from the user by using the
                      'type of alcohol consumed
                      'number of drinks consumed
                      'time in which the drinks were consumed
                      'weight of the user
                      'and finally calculate the BAC of the user
                      
'Purpose of Form : 'have the user click on the time of alcohol was consumed
                   'and arrange that information into arrays
                   'then proceed to the next form
Option Explicit



Private Sub Command1_Click()
If Option1 = True Then
    T = 1
ElseIf Option2 = True Then
    T = 2
ElseIf Option3 = True Then
    T = 3
ElseIf Option4 = True Then
    T = 4
ElseIf Option5 = True Then
    T = 5
End If

Form3.Hide
Form4.Show

End Sub

Private Sub Form_Load()
ReDim Time(1 To 5) As Double


'Open the data file "Time.txt " for the Array

Open "N:\CS130\handin\Sangworn, Jason\Time.txt" For Input As #1
For T = 1 To 5
Input #1, Time(T) 'Read data into the respective Arrays.
Next T
Close #1 'Close data file.


End Sub
