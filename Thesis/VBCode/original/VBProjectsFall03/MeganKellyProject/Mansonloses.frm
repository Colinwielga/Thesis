VERSION 5.00
Begin VB.Form Mansonloses 
   BackColor       =   &H00FF8080&
   Caption         =   "Hey kiddies!"
   ClientHeight    =   7065
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4.906
   ScaleMode       =   5  'Inch
   ScaleWidth      =   4.021
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton QuitMansonL 
      Caption         =   "Quit."
      Height          =   855
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton Quit 
      BackColor       =   &H00FFFF00&
      Caption         =   "Bravely fled Sir Robin..."
      Height          =   975
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9600
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Ok so good news:  You won.  Bad news?  You pissed off a psychotic murderer.  You might want to leave now."
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   5160
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   4785
      Left            =   360
      Picture         =   "Mansonloses.frx":0000
      Top             =   360
      Width           =   5010
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Programmed by Megan Kelly"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Whoah, ok, so good new, you won.  Bad news, you pissed off a homicidal maniac.  So yeah."
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   9120
      Width           =   7695
   End
End
Attribute VB_Name = "Mansonloses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub QuitMansonL_Click()
End
End Sub
