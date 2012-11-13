VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "Type of Alcohol"
   ClientHeight    =   9435
   ClientLeft      =   225
   ClientTop       =   1200
   ClientWidth     =   11880
   FillColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleMode       =   0  'User
   ScaleWidth      =   1.33333e5
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Height          =   3615
      Left            =   8880
      ScaleHeight     =   3555
      ScaleWidth      =   2595
      TabIndex        =   8
      Top             =   3240
      Width           =   2655
   End
   Begin VB.PictureBox Picture2 
      Height          =   3015
      Left            =   5040
      ScaleHeight     =   2955
      ScaleWidth      =   2595
      TabIndex        =   7
      Top             =   3600
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   960
      ScaleHeight     =   3915
      ScaleWidth      =   3075
      TabIndex        =   6
      Top             =   3240
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Choose One of the Options then Click me"
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
      Left            =   4800
      TabIndex        =   4
      Top             =   7560
      Width           =   2895
   End
   Begin VB.OptionButton optwine 
      BackColor       =   &H000000C0&
      Caption         =   "   Wine"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.OptionButton optalcohol 
      BackColor       =   &H00008000&
      Caption         =   " Hard Alcohol"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   5280
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.OptionButton optbeer 
      BackColor       =   &H00400000&
      Caption         =   "   Beer"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   1800
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Left            =   4440
      TabIndex        =   5
      Top             =   1080
      Width           =   3555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "What kind of alcohol have you consumed?"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   3720
      TabIndex        =   3
      Top             =   480
      Width           =   5310
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Blood Alcohol Content (Bloodalcohollevel.vbp)
'Form Name : Blood Alcohol Content (Alcohol.frm)
'Author : Jason Sangworn
'Date Written : March 15, 2004
'Purpose of Project : To calculate the blood alcohol content
                      'from the user by using the
                      'type of alcohol consumed
                      'number of drinks consumed
                      'time in which the drinks were consumed
                      'weight of the user
                      'and finally calculate the BAC of the user
                      
'Purpose of Form : 'have the user click on the type of alcohol consumed
                   'and arrange that information into arrays
                   'then proceed to the next form
Option Explicit

Private Sub Form_Load()
ReDim Alcohol(1 To 3) As Double

Picture2.Picture = LoadPicture("N:\CS130\handin\Sangworn, Jason\cocktail.bmp")
Picture1.Picture = LoadPicture("N:\CS130\handin\Sangworn, Jason\beerbottle.bmp")
Picture3.Picture = LoadPicture("N:\CS130\handin\Sangworn, Jason\wine.bmp")

'Open the data file "alcohol.txt " for the Arrays that

Open "N:\CS130\handin\Sangworn, Jason\alcohol.txt" For Input As #1
For A = 1 To 3
Input #1, Alcohol(A) 'Read data into the respective Arrays.
Next A

Close #1 'Close data file.

End Sub


Private Sub Command1_Click()

If optbeer = True Then
        A = 1
ElseIf optalcohol = True Then
        A = 2
ElseIf optwine = True Then
        A = 3
End If

Form1.Hide
Form2.Show



End Sub



