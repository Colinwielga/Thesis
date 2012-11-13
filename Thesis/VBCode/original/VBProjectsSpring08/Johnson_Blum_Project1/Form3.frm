VERSION 5.00
Begin VB.Form Duluth 
   BackColor       =   &H00800000&
   Caption         =   "Form3"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form3"
   ScaleHeight     =   8640
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      Height          =   2655
      Index           =   0
      Left            =   4320
      ScaleHeight     =   2595
      ScaleWidth      =   4635
      TabIndex        =   6
      Top             =   3840
      Width           =   4695
   End
   Begin VB.CommandButton cmdPicture 
      BackColor       =   &H00808080&
      Caption         =   "Take a glimpse of the Aerial Bridge ==>"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   3375
   End
   Begin VB.CommandButton cmdHomepage 
      BackColor       =   &H8000000D&
      Caption         =   "Back to Homepage"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdTrivia 
      Caption         =   "Answer a Trivia Question"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      TabIndex        =   3
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton cmdFacts 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Press Here For a Fun Fact!"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      MaskColor       =   &H00404040&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.PictureBox picResults 
      Height          =   2775
      Index           =   1
      Left            =   3720
      ScaleHeight     =   2715
      ScaleWidth      =   3795
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lblDuluth 
      Caption         =   "       Duluth"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro B"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Duluth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFacts_Click()
picResults.Cls



End Sub

Private Sub cmdHomepage_Click()
Duluth.Hide
Minnesota.Show

End Sub

Private Sub cmdPicture_Click()
picResults2.Print LoadPicture(App.Path & "\duluth_bridge_edit.jpg")

End Sub

Private Sub cmdTrivia_Click()
Dim Boat As String
Boat = InputBox("What is the name of the infamous ship that sank in Lake Superior in 1975? (make sure you get the spelling right ;)")

If Boat = "Edmund Fitzgerald" Then
    MsgBox ("You are correct! " & Boat & " was the name of the ship that sank...what a tradegy.")
    
Else
    MsgBox ("No, I'm sorry, the name was Edmund Fitzgerald.")
    
    End If
End Sub
