VERSION 5.00
Begin VB.Form frmbottoms 
   BackColor       =   &H00FF8080&
   Caption         =   "Bottoms"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdHome 
      Caption         =   "Go Home!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2760
      Picture         =   "frmbottoms.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdChoose2 
      Caption         =   "Choose!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3060
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtbottom 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   1
      Top             =   750
      Width           =   1215
   End
   Begin VB.Label lblcolor 
      Caption         =   " Choose from:   1. Black   2. Purple     3. Red    4. Blue"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   870
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a bottom color!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   30
      Width           =   3495
   End
End
Attribute VB_Name = "frmbottoms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChoose2_Click()
'opens an array and compares user input with the array to pick a color that will be printed in a picResults box
btmcolor = txtbottom.Text
Dim CTR As Integer
Dim colors(1 To 100) As String
Open App.Path & "\colors.txt" For Input As #1
CTR = 0
Do Until EOF(1)
CTR = CTR + 1
Input #1, colors(CTR)
Loop
Close #1
picResults.Print colors(btmcolor)
End Sub

Private Sub cmdHome_Click()
'opens a new form
frmDoll.Show
frmbottoms.Hide
End Sub

