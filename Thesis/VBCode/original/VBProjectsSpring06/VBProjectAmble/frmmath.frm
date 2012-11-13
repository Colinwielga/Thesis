VERSION 5.00
Begin VB.Form frmmath 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6300
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   975
   End
   Begin VB.PictureBox picresults4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   2760
      Width           =   615
   End
   Begin VB.PictureBox picresults3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   2760
      Width           =   615
   End
   Begin VB.PictureBox picresults2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   2760
      Width           =   615
   End
   Begin VB.PictureBox picresults1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox txtnumber 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmdrandom 
      Caption         =   "Enter"
      Height          =   615
      Left            =   5160
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lbljeff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Created By: Jeff Amble"
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Image imgclover 
      Height          =   1710
      Left            =   4200
      Picture         =   "frmmath.frx":0000
      Top             =   960
      Width           =   1725
   End
   Begin VB.Label lblnumbers 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your Lucky Numbers are:"
      BeginProperty Font 
         Name            =   "MS Mincho"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label lblexplain 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter your age here and get your lucky number for the day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "frmmath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form enables the user to enter their age and pulse'
'and get their heart rate back and whether or not they are'
'healthy'
Option Explicit
Dim A As Integer
Dim B As Integer
Dim C As Integer
Dim D As Integer
Dim E As Integer
Dim X As Integer
'This button enables the user to go back to the main form'
Private Sub cmdback_Click()
    frmmath.Visible = False
    frmmain.Visible = True
End Sub
'This button gives the user their lucky numbers'
Private Sub cmdrandom_Click()
    picresults.Cls
    picresults1.Cls
    picresults2.Cls
    picresults3.Cls
    picresults4.Cls
    X = txtnumber
        A = Int(Int((13 - 3 / 2 + 9) * Rnd * (X / 9)))
            picresults.Print A
        B = Int(Int((8 - 4 / 3 + 13) * Rnd * (X / 13)))
            picresults1.Print B
        C = Int(Int((15 - 3 / 5 + 11) * Rnd * (X / 11)))
            picresults2.Print C
        D = Int(Int((5 - 1 / 3 + 10) * Rnd * (X / 10)))
            picresults3.Print D
        E = Int(Int((3 - 5 / 3 + 7) * Rnd * (X / 7)))
            picresults4.Print E
End Sub

'All five picboxes display the users' lucky numbers'
Private Sub picresults_Click()

End Sub
