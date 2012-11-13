VERSION 5.00
Begin VB.Form frmCops 
   BackColor       =   &H00404040&
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      Height          =   4455
      Left            =   11640
      Picture         =   "Cops.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   3195
      TabIndex        =   24
      Top             =   2400
      Width           =   3255
   End
   Begin VB.PictureBox Picture4 
      Height          =   975
      Left            =   840
      Picture         =   "Cops.frx":367AA
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   23
      Top             =   8040
      Width           =   1335
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   960
      Picture         =   "Cops.frx":3A9CC
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   22
      Top             =   4920
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   960
      Picture         =   "Cops.frx":3EBEE
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   21
      Top             =   1920
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   4920
      Picture         =   "Cops.frx":42E10
      ScaleHeight     =   2475
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fox        Ch.9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   20
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fox        Ch.9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   19
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fox        Ch.9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   18
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fox        Ch.9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   17
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fox        Ch.9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   16
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5040
      TabIndex        =   15
      Top             =   6120
      Width           =   5655
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5040
      TabIndex        =   14
      Top             =   6840
      Width           =   5655
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5040
      TabIndex        =   13
      Top             =   5400
      Width           =   5655
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5040
      TabIndex        =   12
      Top             =   3960
      Width           =   5655
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5040
      TabIndex        =   11
      Top             =   4680
      Width           =   5655
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2:00PM -- 2:30PM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1:30PM -- 2:00PM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   9
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1:00PM -- 1:30PM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12:30PM -- 1:00PM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12:00PM -- 12:30PM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Thursday"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Friday"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Wednesday"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tuesday"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Monday"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu Return 
         Caption         =   "Return to Menu"
      End
      Begin VB.Menu quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmCops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Crime Awareness Project
'frm License
'By: Alex Warzecha & Andrew Enzler
'Written: 3/18/2009
'Objective: to offer the schedule for the popular criminal television show COPS.


Private Sub quit_Click()
End
End Sub

Private Sub return_Click()
frmCops.Hide
frmHome.Show
End Sub
