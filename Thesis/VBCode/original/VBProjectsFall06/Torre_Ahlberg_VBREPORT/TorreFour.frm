VERSION 5.00
Begin VB.Form TorreFour 
   BackColor       =   &H0080FF80&
   Caption         =   "Recieving Your Computer ID"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous Page"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.PictureBox picSwimmingWorld 
      Height          =   1215
      Left            =   2640
      Picture         =   "TorreFour.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   5475
      TabIndex        =   4
      Top             =   0
      Width           =   5535
   End
   Begin VB.PictureBox picKidComputer 
      Height          =   3135
      Left            =   4320
      Picture         =   "TorreFour.frx":C28F
      ScaleHeight     =   3075
      ScaleWidth      =   3795
      TabIndex        =   3
      Top             =   1920
      Width           =   3855
   End
   Begin VB.PictureBox picComputer 
      Height          =   2415
      Left            =   0
      Picture         =   "TorreFour.frx":EE98
      ScaleHeight     =   2355
      ScaleWidth      =   2115
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
   Begin VB.PictureBox picOutput 
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
   End
   Begin VB.CommandButton cmdFindID 
      Caption         =   "Find Coumputer ID"
      Height          =   975
      Left            =   2280
      TabIndex        =   0
      Top             =   2520
      Width           =   1815
   End
End
Attribute VB_Name = "TorreFour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Author of this form is Torre Ahlberg on 11/02/06
'The Purpose of this form is to countinue and finish what was started on the previous form
'it has the user enter their name and get a computer ID name for a swimming website

'This Subroutine utilizes String Functions and the name input in the Inputbox on the previous page
'to create a computer ID for a swimming website
Private Sub cmdFindID_Click()
    picOutput.Cls
    N = InStr(YourName, "")
    First = Left(YourName, N - 1)
    Last = Right(YourName, Len(YourName) - (N + 2))
    Middle = Mid(YourName, N + 1, 1)
    ID = Left(First, 1) & Middle & Left(Last, 6)
    picOutput.Print "Your User Name Is", ID

End Sub
'I borrowed some of this code from our textbook on Page[7]11
'I hope this is the citing you were looking for

Private Sub cmdPrevious_Click()
    TorreFour.Visible = False
    TorreThree.Visible = True
End Sub
