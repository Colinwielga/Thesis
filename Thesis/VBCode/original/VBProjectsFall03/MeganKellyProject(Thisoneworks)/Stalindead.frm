VERSION 5.00
Begin VB.Form stalindead 
   BackColor       =   &H00000000&
   Caption         =   "Ding, dong, Stalin's dead!"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form3"
   ScaleHeight     =   6540
   ScaleWidth      =   7095
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton quitstalin 
      BackColor       =   &H80000014&
      Caption         =   "Quit"
      Height          =   1095
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Programmed by Megan Kelly"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Poor Stalin.... it's so very very sad when evil, murderous dictators get what's coming to them...."
      ForeColor       =   &H80000014&
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   3840
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   3345
      Left            =   1080
      Picture         =   "Stalindead.frx":0000
      Top             =   480
      Width           =   4425
   End
End
Attribute VB_Name = "stalindead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub quitstalin_Click()
End
End Sub
