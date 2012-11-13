VERSION 5.00
Begin VB.Form frmdisc 
   BackColor       =   &H80000013&
   Caption         =   "Discography"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   4680
   FillColor       =   &H00FF0000&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdhard 
      Caption         =   "Click to see songs"
      Height          =   375
      Index           =   5
      Left            =   3840
      TabIndex        =   17
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmddesert 
      Caption         =   "Click to see songs"
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   16
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdrecovering 
      Caption         =   "Click to see songs"
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   15
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdaugust 
      Caption         =   "Click to see songs"
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   14
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdghosts 
      Caption         =   "Click to see songs"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   13
      Top             =   2760
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   3480
      Picture         =   "frmdisc.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   3600
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Index           =   2
      Left            =   600
      Picture         =   "frmdisc.frx":13BD2
      ScaleHeight     =   2115
      ScaleWidth      =   2235
      TabIndex        =   5
      Top             =   3600
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Index           =   5
      Left            =   3480
      Picture         =   "frmdisc.frx":25554
      ScaleHeight     =   2115
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      FillColor       =   &H0000FF00&
      ForeColor       =   &H8000000D&
      Height          =   3735
      Left            =   600
      ScaleHeight     =   3675
      ScaleWidth      =   3675
      TabIndex        =   3
      Top             =   6600
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Index           =   1
      Left            =   6360
      Picture         =   "frmdisc.frx":360C6
      ScaleHeight     =   2115
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Index           =   0
      Left            =   600
      Picture         =   "frmdisc.frx":49818
      ScaleHeight     =   2115
      ScaleWidth      =   2235
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H000000FF&
      Caption         =   "Go back to the Main Page"
      Height          =   735
      Left            =   9360
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Created by Matt Proulx"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   18
      Top             =   10440
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Songs"
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000013&
      Caption         =   "August and Everything After- 1993"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000013&
      Caption         =   "Recovering the Satalites-  1996"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "The Desert Life- 1999 "
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "Hard Candy- 2002"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "Films About Ghosts- 2003"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   7
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmdisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : CountingCrows (Matt Proulx's VB Project.vbp)
'Form Name : frmdisc (frmdisc.frm)
'Author: Matt Proulx
'Date Written: March 13, 2004
'Purpose of the Form: 'This form lets the user view all of the Counting Crows Cd's. It will also let the user view the
                      'songs from each cd by clicking on the appropriate button under each cd.

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Dim Path As String
Dim Songs As String
Dim CTR As Single
Private Sub cmdaugust_Click(Index As Integer)
picResults.Cls 'Clears the box
Open Path & "august.txt" For Input As #1 'Opens the notepad with songs on it
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, Songs
        picResults.Print Songs
    Loop
    Close #1
End Sub
Private Sub cmdback_Click()
    frmdisc.Hide
    frmtitle.Show
End Sub
Private Sub cmddesert_Click(Index As Integer)
picResults.Cls
Open Path & "desert.txt" For Input As #1 'Opens the notepad with songs on it
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, Songs
        picResults.Print Songs 'Prints the songs in the display window
    Loop
    Close #1
End Sub
Private Sub cmdghosts_Click(Index As Integer)
picResults.Cls
Open Path & "films.txt" For Input As #1 'Opens the notepad with songs on it
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, Songs
        picResults.Print Songs 'Prints the songs in the display window
    Loop
    Close #1
End Sub
Private Sub cmdhard_Click(Index As Integer)
picResults.Cls
Open Path & "hard.txt" For Input As #1 'Opens the notepad with songs on it
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, Songs
        picResults.Print Songs 'Prints the songs in the display window
    Loop
    Close #1
End Sub
Private Sub cmdrecovering_Click(Index As Integer)
picResults.Cls
Open Path & "recovering.txt" For Input As #1 'Opens the notepad with songs on it
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, Songs
        picResults.Print Songs 'Prints the songs in the display window
    Loop
    Close #1
End Sub
Private Sub Form_Load()
    Path = "N:\CS130\handin\Proulx, Matt\"
End Sub


