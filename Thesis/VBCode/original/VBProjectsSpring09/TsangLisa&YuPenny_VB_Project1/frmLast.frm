VERSION 5.00
Begin VB.Form frmLast 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Final Step"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11295
   FillStyle       =   4  'Upward Diagonal
   BeginProperty Font 
      Name            =   "Rockwell"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLast.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmLast.frx":030A
   ScaleHeight     =   7260
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAnswer 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   4440
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Frame fraQuestion 
      BackColor       =   &H80000009&
      Caption         =   "Do you find this program useful ? Please Enter ""Yes"" or ""No"". Thank you!"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   1440
      TabIndex        =   3
      Top             =   2760
      Width           =   8415
      Begin VB.CommandButton cmdSubmit 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Submit"
         Height          =   495
         Left            =   5760
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Shape spe1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4080
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblEnjoy 
      BackColor       =   &H80000009&
      Caption         =   "Enjoy your grocery shopping and cooking them..!"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   7455
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H80000009&
      Caption         =   "Thank you for using our program for planning your meal. "
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   10335
   End
End
Attribute VB_Name = "frmLast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Answer As String, Yes As String, No As String


Private Sub cmdSubmit_Click()

Answer = txtAnswer.Text

If Answer = "Yes" Then
    MsgBox ("Thank you. And have a great day.")
    End
ElseIf Answer = "No" Then
    MsgBox ("Thank you for your opinion. We will work harder next time.")
    End
Else
    MsgBox ("Please type in Yes or No.")
End If

End Sub
