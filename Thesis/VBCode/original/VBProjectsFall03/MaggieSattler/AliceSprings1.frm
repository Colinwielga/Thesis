VERSION 5.00
Begin VB.Form AliceSprings1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   10785
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
      Height          =   1095
      Left            =   5520
      TabIndex        =   3
      Top             =   5520
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
      Height          =   1095
      Left            =   720
      TabIndex        =   2
      Top             =   5520
      Width           =   3855
   End
   Begin VB.PictureBox Imagebox 
      Height          =   3135
      Left            =   2400
      ScaleHeight     =   3075
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Maggie Sattler"
      Height          =   255
      Left            =   7920
      TabIndex        =   4
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Alice Springs Images"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "AliceSprings1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim P As Integer



Private Sub Next_Click()

P = P + 1
If P = 5 Then P = 1
If P = 1 Then
        Imagebox.Picture = LoadPicture(Australia1.PATH & "AliceSprings1.jpg")
    ElseIf P = 2 Then
        Imagebox.Picture = LoadPicture(Australia1.PATH & "AliceSprings2.jpg")
    ElseIf P = 3 Then
        Imagebox.Picture = LoadPicture(Australia1.PATH & "AliceSprings3.jpg")
    ElseIf P = 4 Then
        Imagebox.Picture = LoadPicture(Australia1.PATH & "AliceSprings4.jpg")
End If
        
End Sub

Private Sub Quit_Click()
    AliceSprings1.Hide
    Australia1.Show
End Sub
