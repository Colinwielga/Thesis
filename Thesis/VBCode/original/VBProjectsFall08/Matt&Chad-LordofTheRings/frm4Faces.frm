VERSION 5.00
Begin VB.Form frm4Faces 
   Caption         =   "Form2"
   ClientHeight    =   7440
   ClientLeft      =   1635
   ClientTop       =   1905
   ClientWidth     =   9915
   LinkTopic       =   "Form2"
   Picture         =   "frm4Faces.frx":0000
   ScaleHeight     =   7440
   ScaleWidth      =   9915
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Your Journey"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back"
      Height          =   255
      Left            =   8520
      TabIndex        =   10
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdFrodo 
      Caption         =   "Frodo"
      Height          =   615
      Left            =   1080
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdBoromir 
      Caption         =   "Boromir"
      Height          =   615
      Left            =   8280
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAragorn 
      Caption         =   "Aragorn"
      Height          =   615
      Left            =   6960
      TabIndex        =   7
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdMerry 
      Caption         =   "Merry"
      Height          =   615
      Left            =   1800
      TabIndex        =   6
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdSam 
      Caption         =   "Sam"
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdGimli 
      Caption         =   "Gimli"
      Height          =   615
      Left            =   8280
      TabIndex        =   4
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdLegolas 
      Caption         =   "Legolas"
      Height          =   615
      Left            =   6960
      TabIndex        =   3
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdGandalf 
      Caption         =   "Gandalf"
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdPippin 
      Caption         =   "Pippin"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   5880
      Width           =   1095
   End
   Begin VB.PictureBox picture1 
      Height          =   2655
      Left            =   3120
      ScaleHeight     =   2595
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   3840
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Click the name of the character you wish to see!"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1095
      Left            =   480
      TabIndex        =   12
      Top             =   240
      Width           =   9375
   End
End
Attribute VB_Name = "frm4Faces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAragorn_Click()
    picture1.Picture = LoadPicture(App.Path & "\Aragorn.jpg")
End Sub

Private Sub cmdBoromir_Click()
    picture1.Picture = LoadPicture(App.Path & "\Boromir.jpg")
End Sub

Private Sub cmdGandalf_Click()
    picture1.Picture = LoadPicture(App.Path & "\Gandalf.jpg")
End Sub

Private Sub cmdGimli_Click()
    picture1.Picture = LoadPicture(App.Path & "\Gimli.jpg")
End Sub

Private Sub cmdGoBack_Click()
    frm4Faces.Hide
    frm2Characters.Show
End Sub

Private Sub cmdLegolas_Click()
    picture1.Picture = LoadPicture(App.Path & "\Legolas.jpg")
End Sub

Private Sub cmdMerry_Click()
    picture1.Picture = LoadPicture(App.Path & "\Merry.jpg")
End Sub

Private Sub cmdPippin_Click()
    picture1.Picture = LoadPicture(App.Path & "\Pippin.jpg")
End Sub

Private Sub cmdQuit_Click()
    Quit
End Sub
Private Sub cmdFrodo_Click()
    picture1.Picture = LoadPicture(App.Path & "\Frodo.jpg")
End Sub
Private Sub cmdSam_Click()
    picture1.Picture = LoadPicture(App.Path & "\Sam.jpg")
End Sub

Private Sub Form_Activate()
    picture1.Picture = LoadPicture("") 'this clears the picture box when form is loaded
End Sub
