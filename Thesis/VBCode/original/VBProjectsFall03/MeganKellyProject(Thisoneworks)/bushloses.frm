VERSION 5.00
Begin VB.Form beatbush 
   BackColor       =   &H000000FF&
   Caption         =   "He's so endearingly pathetic..."
   ClientHeight    =   7215
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   7470
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   2280
      TabIndex        =   1
      Top             =   8520
      Width           =   2295
   End
   Begin VB.CommandButton QuitBushL 
      Caption         =   "Quit"
      Height          =   855
      Left            =   2280
      TabIndex        =   0
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Uh oh!  You made Georgie lose!  You better get out of here before he calls his Daddy!"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   4560
      Width           =   5535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Programmed by Megan Kelly"
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   4290
      Left            =   720
      Picture         =   "bushloses.frx":0000
      Top             =   120
      Width           =   6150
   End
End
Attribute VB_Name = "beatbush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub QuitBushL_Click()
End
End Sub
