VERSION 5.00
Begin VB.Form frmshirts 
   BackColor       =   &H00FF80FF&
   Caption         =   "Shirts"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Go home!"
      Height          =   2175
      Left            =   4680
      Picture         =   "frmshirts.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdChoose 
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
      Height          =   975
      Left            =   5040
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.PictureBox picResults2 
      Height          =   735
      Left            =   1200
      ScaleHeight     =   675
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txtColor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "   1. Black     2. Purple    3. Red     "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Label lblColor 
      Caption         =   "Enter Shirt Color:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "frmshirts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChoose_Click()
'Reads an Array and prints the color chosen by the user
Dim CTR As Integer
Dim colors(1 To 100) As String
Open App.Path & "\colors.txt" For Input As #1
CTR = 0
Do Until EOF(1)
CTR = CTR + 1
Input #1, colors(CTR)
Loop
Close #1
shcolor = txtColor.Text
picResults2.Print colors(shcolor)

End Sub

Private Sub cmdReturn_Click()
'Opens new form
frmshirts.Hide
frmDoll.Show
End Sub
