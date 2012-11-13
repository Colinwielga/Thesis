VERSION 5.00
Begin VB.Form Broome1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Left            =   4440
      TabIndex        =   3
      Top             =   5520
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
      Left            =   360
      TabIndex        =   2
      Top             =   5520
      Width           =   3375
   End
   Begin VB.PictureBox Imagebox 
      Height          =   3735
      Left            =   960
      ScaleHeight     =   3675
      ScaleWidth      =   6075
      TabIndex        =   1
      Top             =   1440
      Width           =   6135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Maggie Sattler"
      Height          =   255
      Left            =   6600
      TabIndex        =   4
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Broome Images"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
End
Attribute VB_Name = "Broome1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim P As Integer




Private Sub Command1_Click()
    Broome1.Hide
    Australia1.Show
    
End Sub

Private Sub Next_Click()

P = P + 1
If P = 5 Then P = 1
If P = 1 Then
        Imagebox.Picture = LoadPicture(Australia1.PATH & "Broome1.jpg")
    ElseIf P = 2 Then
        Imagebox.Picture = LoadPicture(Australia1.PATH & "Broome2.jpg")
    ElseIf P = 3 Then
        Imagebox.Picture = LoadPicture(Australia1.PATH & "Broome3.jpg")
    ElseIf P = 4 Then
        Imagebox.Picture = LoadPicture(Australia1.PATH & "Broome4.jpg")
End If

End Sub
