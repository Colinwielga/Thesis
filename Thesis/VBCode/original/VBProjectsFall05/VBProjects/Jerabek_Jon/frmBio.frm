VERSION 5.00
Begin VB.Form frmBio 
   BackColor       =   &H00800000&
   Caption         =   "Biography"
   ClientHeight    =   4845
   ClientLeft      =   3945
   ClientTop       =   3285
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   7800
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   6480
      Picture         =   "frmBio.frx":0000
      ScaleHeight     =   1575
      ScaleWidth      =   1215
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO!"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.PictureBox picOutput 
      Height          =   615
      Left            =   4080
      ScaleHeight     =   555
      ScaleWidth      =   2835
      TabIndex        =   7
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox txtAge 
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.PictureBox picKGYoung 
      Height          =   2775
      Left            =   120
      Picture         =   "frmBio.frx":21F2
      ScaleHeight     =   2715
      ScaleWidth      =   2475
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmBio.frx":305B4
      Top             =   3600
      Width           =   5535
   End
   Begin VB.TextBox txtBio 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmBio.frx":307AE
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton cmdMain2 
      Caption         =   "Main"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   0
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label lblCompare 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Compare your age to KG's"
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblInteresting 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "INTERESTING FACTS"
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   1935
   End
End
Attribute VB_Name = "frmBio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ProjectKG
'frmBio
'Jon Jerabek
'10-25-05 & 10-26-05
'Objective-User can compare their age to KG's. Also view interesting facts

Private Sub cmdGo_Click()    'User inputs age. Age is compared to Kevin's. Appropriate outcome is displayed.
picOutput.Cls
Dim x As Double
Dim y As Double
Dim Sum As Double
x = txtAge.Text
y = 29
Select Case x
Case Is > y
    Sum = (x - y)
    picOutput.Print "You are"; " "; Sum; " "; "years older than KG."
Case Is < y
    Sum = (y - x) + Sum
    picOutput.Print "You are"; " "; Sum; " "; "years younger than KG."
Case Is = y
    picOutput.Print "You are the same age as KG!"
Case Else
    picOutput.Print "Error"
End Select

    
End Sub

Private Sub cmdMain2_Click()
frmHome.Show
frmBio.Hide
End Sub

