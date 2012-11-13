VERSION 5.00
Begin VB.Form Perth1 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   10725
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
      Height          =   1095
      Left            =   6240
      TabIndex        =   3
      Top             =   5880
      Width           =   3255
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
      Height          =   1095
      Left            =   1200
      TabIndex        =   2
      Top             =   5880
      Width           =   3375
   End
   Begin VB.PictureBox ImageBox 
      Height          =   3735
      Left            =   2880
      ScaleHeight     =   3675
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Maggie Sattler"
      Height          =   255
      Left            =   8160
      TabIndex        =   4
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Perth Images"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "Perth1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim P As Integer





Private Sub Command2_Click()
    Perth1.Hide
    Australia1.Show
End Sub

Private Sub Next_Click()


    P = P + 1
    If P = 5 Then P = 1
    If P = 1 Then
            Imagebox.Picture = LoadPicture(Australia1.PATH & "Perth1.jpg")
        ElseIf P = 2 Then
            Imagebox.Picture = LoadPicture(Australia1.PATH & "Perth2.jpg")
        ElseIf P = 3 Then
            Imagebox.Picture = LoadPicture(Australia1.PATH & "Perth3.jpg")
        ElseIf P = 4 Then
            Imagebox.Picture = LoadPicture(Australia1.PATH & "Perth4.jpg")
    End If



End Sub
