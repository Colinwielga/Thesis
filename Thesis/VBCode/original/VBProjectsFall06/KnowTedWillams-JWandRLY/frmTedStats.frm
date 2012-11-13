VERSION 5.00
Begin VB.Form frmTedStats 
   Caption         =   "Ted's Statistics"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   Picture         =   "frmTedStats.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInput 
      BackColor       =   &H000000FF&
      Caption         =   "Click to Input  Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   4455
   End
   Begin VB.PictureBox picResults 
      Height          =   2415
      Left            =   240
      ScaleHeight     =   2355
      ScaleWidth      =   14715
      TabIndex        =   1
      Top             =   240
      Width           =   14775
   End
   Begin VB.CommandButton cmdTedmenuback 
      BackColor       =   &H000000FF&
      Caption         =   "Click to Return to Menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9000
      Width           =   4455
   End
End
Attribute VB_Name = "frmTedStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdInput_Click()
    Dim Year As Integer, Stats() As String, Found As Boolean, ctr As Single
    Found = False
        Open App.Path & "\TedYearlyStats.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Stats()
        Loop
    Year = InputBox("Enter a Year of Ted's Career, 1939-1960", "Enter Year 1939-1960")
        Do While (Not Found) And Year >= 1939
            ctr = ctr + 1
            If Year = Stats(ctr) Then Found = True
        Loop
End Sub

Private Sub cmdTedmenuBack_Click()
    frmTedmenu.Show
    frmTedStats.Hide
End Sub

