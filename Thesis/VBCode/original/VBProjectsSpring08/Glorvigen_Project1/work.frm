VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form14"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   LinkTopic       =   "Form14"
   ScaleHeight     =   4845
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdwork 
      BackColor       =   &H0080C0FF&
      Caption         =   "Work's Cited"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   360
      ScaleHeight     =   4395
      ScaleWidth      =   7155
      TabIndex        =   2
      Top             =   240
      Width           =   7215
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H008080FF&
      Caption         =   "Leave Minnesota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Back To Main Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdback_Click()
    Form14.Hide
    form1.Show
End Sub

Private Sub cmdexit_Click()
    End
End Sub

Private Sub cmdwork_Click()
Dim work(1 To 14) As String
Dim ctr As Integer
Dim n As Integer

ctr = 0
    
    picoutput.Cls
    picoutput.Print "Work's Cited"
    picoutput.Print "*******************************************************"
    
     Open App.Path & "\work.txt" For Input As #1
            Do Until EOF(1)
                ctr = ctr + 1
                Input #1, work(ctr)
            Loop
        Close #1
        
    For n = 1 To 13
        picoutput.Print work(n)
    Next n
    
    
End Sub
