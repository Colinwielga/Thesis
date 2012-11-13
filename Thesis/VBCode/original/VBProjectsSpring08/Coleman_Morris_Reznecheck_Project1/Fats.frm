VERSION 5.00
Begin VB.Form frm8 
   BackColor       =   &H00004080&
   Caption         =   "Oils"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picoutput 
      Height          =   4575
      Left            =   2640
      ScaleHeight     =   4515
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "Show Examples"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back To Food Pyramid"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004080&
      Caption         =   $"Fats.frx":0000
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frm8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Healthy Living
'frm8
'Ben Morris
'March 21
'displays the different fats
Dim oil(1 To 7) As String
Dim CTR As Integer
Option Explicit

Private Sub cmdback_Click()
    frm1.Show
    frm2.Hide
    frm3.Hide
    frm4.Hide
    frm5.Hide
    frm6.Hide
    frm7.Hide
    frm8.Hide
End Sub

Private Sub cmdshow_Click()
    
     picoutput.Cls
    CTR = 0
    picoutput.Print "These are some Examples of Oils"
    picoutput.Print "--------------------------------------------------------------------"
    Open App.Path & "\oil.txt" For Input As #1
    Do Until EOF(1)
    
    ' Get the data from the file
        CTR = CTR + 1
        Input #1, oil(CTR)
        
        picoutput.Print oil(CTR)
    
    Loop
    Close #1
    
End Sub

