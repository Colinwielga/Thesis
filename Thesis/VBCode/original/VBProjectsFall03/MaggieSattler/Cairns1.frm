VERSION 5.00
Begin VB.Form Cairns1 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Height          =   975
      Left            =   5400
      TabIndex        =   3
      Top             =   5280
      Width           =   3375
   End
   Begin VB.CommandButton Next 
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
      Height          =   975
      Left            =   960
      TabIndex        =   2
      Top             =   5280
      Width           =   3615
   End
   Begin VB.PictureBox Imagebox 
      Height          =   2895
      Left            =   3480
      ScaleHeight     =   2835
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Maggie Sattler"
      Height          =   255
      Left            =   7680
      TabIndex        =   4
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Cairns Images"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "Cairns1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim P As Integer





Private Sub Command2_Click()
    Cairns1.Hide
    Australia1.Show
End Sub

Private Sub Next_Click()

    P = P + 1
    If P = 5 Then P = 1
    If P = 1 Then
            Imagebox.Picture = LoadPicture(Australia1.PATH & "Cairns1.jpg")
        ElseIf P = 2 Then
            Imagebox.Picture = LoadPicture(Australia1.PATH & "Cairns2.jpg")
        ElseIf P = 3 Then
            Imagebox.Picture = LoadPicture(Australia1.PATH & "Cairns3.jpg")
        ElseIf P = 4 Then
            Imagebox.Picture = LoadPicture(Australia1.PATH & "Cairns4.jpg")
    End If
End Sub



