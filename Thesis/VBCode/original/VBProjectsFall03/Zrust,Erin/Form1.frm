VERSION 5.00
Begin VB.Form Dave_Matthews_Band_CDs 
   BackColor       =   &H000040C0&
   Caption         =   "Dave Matthews Band CDs"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   255
      Left            =   7920
      TabIndex        =   15
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   255
      Left            =   5640
      TabIndex        =   14
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   2760
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000C000&
      Caption         =   "Next Band/Songwriter"
      Height          =   735
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Add to Your Shopping Cart"
      Height          =   735
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   2655
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0000C000&
      Height          =   3615
      Left            =   600
      ScaleHeight     =   3555
      ScaleWidth      =   4995
      TabIndex        =   4
      Top             =   3720
      Width           =   5055
   End
   Begin VB.PictureBox Picture4 
      Height          =   1815
      Left            =   7200
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.PictureBox Picture3 
      Height          =   1815
      Left            =   4800
      Picture         =   "Form1.frx":6FE3
      ScaleHeight     =   1755
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Height          =   1935
      Left            =   2520
      Picture         =   "Form1.frx":CA17
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   360
      Picture         =   "Form1.frx":130FE
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H000040C0&
      Caption         =   "Busted Stuff"
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H000040C0&
      Caption         =   "Live in Chicago- United Center"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H000040C0&
      Caption         =   "Before These Crowded Streets"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H000040C0&
      Caption         =   "Crash"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H000040C0&
      Caption         =   "Select a Dave Matthews Band CD"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   3120
      Width           =   2655
   End
End
Attribute VB_Name = "Dave_Matthews_Band_CDs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

End Sub

Private Sub Form_Load()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Picture5_Click()

End Sub
