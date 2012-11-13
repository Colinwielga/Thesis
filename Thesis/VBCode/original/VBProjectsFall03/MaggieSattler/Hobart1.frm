VERSION 5.00
Begin VB.Form Hobart1 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quit 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   3
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton SeePic 
      Caption         =   "Next Image"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   4920
      Width           =   3135
   End
   Begin VB.PictureBox Imagebox 
      Height          =   3495
      Left            =   1320
      ScaleHeight     =   3435
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Maggie Sattler"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Hobart Images"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   735
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Hobart1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim P As Integer



Private Sub Quit_Click()
    Hobart1.Hide
    Australia1.Show
    
End Sub

Private Sub SeePic_Click()

P = P + 1
If P = 5 Then P = 1
If P = 1 Then
        Imagebox.Picture = LoadPicture(Australia1.PATH & "Hobart1.jpg")
    ElseIf P = 2 Then
        Imagebox.Picture = LoadPicture(Australia1.PATH & "Hobart2.jpg")
    ElseIf P = 3 Then
        Imagebox.Picture = LoadPicture(Australia1.PATH & "Hobart3.jpg")
    ElseIf P = 4 Then
        Imagebox.Picture = LoadPicture(Australia1.PATH & "Hobart4.jpg")
End If

End Sub
