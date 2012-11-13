VERSION 5.00
Begin VB.Form frmB 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   Picture         =   "frmB.frx":0000
   ScaleHeight     =   6780
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   2955
      TabIndex        =   4
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eat Lambis Bread on Eve of Battle"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "End Your Journey"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Height          =   255
      Left            =   5400
      TabIndex        =   1
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmB.frx":75E1
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   6975
   End
End
Attribute VB_Name = "frmB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmB.Hide
    frmC.Show
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    If Lambis > 0 Then
        Lambis = Lambis - 1
        picResults.Cls
        picResults.Print "You have "; Lambis; "pieces of Lambis Bread left"
    End If
End Sub
