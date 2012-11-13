VERSION 5.00
Begin VB.Form frm4star 
   BackColor       =   &H80000011&
   Caption         =   "Form2"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14970
   LinkTopic       =   "Form2"
   Picture         =   "frm4star.frx":0000
   ScaleHeight     =   9825
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "I Want To Book This Hotel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8760
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Return to Hotels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8760
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000011&
      Caption         =   $"frm4star.frx":1051D
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   10920
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000011&
      Caption         =   "The Borders Lodge"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   90
         Charset         =   2
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   480
      TabIndex        =   0
      Top             =   6840
      Width           =   15015
   End
End
Attribute VB_Name = "frm4star"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FormHotel.Show
frm4star.Hide
End Sub

Private Sub Command2_Click()
Dim Nights As Integer
Nights = InputBox("How Many Nights Do You Want To Stay")
totalhotelcost = 0
totalhotelcost = 500 * Nights
MsgBox "" & Nights & " nights in The Borders Lodge will cost you " & FormatCurrency(totalhotelcost) & " dollars "
End Sub
