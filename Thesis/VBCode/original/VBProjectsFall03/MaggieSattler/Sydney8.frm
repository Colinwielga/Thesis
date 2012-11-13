VERSION 5.00
Begin VB.Form Sydney1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   8355
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
      Height          =   855
      Left            =   4800
      TabIndex        =   3
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
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
      Left            =   360
      TabIndex        =   2
      Top             =   5040
      Width           =   3015
   End
   Begin VB.PictureBox Imagebox 
      Height          =   3375
      Left            =   1560
      ScaleHeight     =   3315
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Maggie Sattler"
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sydney Images"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Sydney1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim P As Integer





Private Sub Command1_Click()
    
    P = P + 1
    If P = 7 Then P = 1
    If P = 1 Then
            Imagebox.Picture = LoadPicture(Australia1.PATH & "Sydney1.jpg")
        ElseIf P = 2 Then
            Imagebox.Picture = LoadPicture(Australia1.PATH & "Sydney2.jpg")
        ElseIf P = 3 Then
            Imagebox.Picture = LoadPicture(Australia1.PATH & "Sydney3.jpg")
        ElseIf P = 4 Then
            Imagebox.Picture = LoadPicture(Australia1.PATH & "Sydney4.jpg")
        ElseIf P = 5 Then
            Imagebox.Picture = LoadPicture(Australia1.PATH & "Sydney5.jpg")
        ElseIf P = 6 Then
            Imagebox.Picture = LoadPicture(Australia1.PATH & "Sydney6.jpg")
            
    End If
End Sub

Private Sub Command2_Click()
    Sydney1.Hide
    Australia1.Show
    
End Sub

