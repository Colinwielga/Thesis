VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1500button 
      Caption         =   "1500 Meters"
      Height          =   975
      Left            =   3840
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command10Kbutton 
      Caption         =   "10K"
      Height          =   975
      Left            =   3840
      TabIndex        =   5
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Comman5Kbutton 
      Caption         =   "5K"
      Height          =   735
      Left            =   3840
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Steeplebutton 
      Caption         =   "Steeplechase"
      Height          =   735
      Left            =   3840
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Qutibutton 
      Caption         =   "Quit"
      Height          =   975
      Left            =   6960
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "PLease Select A RacE:"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Outdoor Track and Field Performance Program"
      BeginProperty Font 
         Name            =   "BatangChe"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1500button_Click()

End Sub

Private Sub Qutibutton_Click()
End
End Sub
