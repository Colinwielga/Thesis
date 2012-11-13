VERSION 5.00
Begin VB.Form frm5 
   BackColor       =   &H00404080&
   Caption         =   "Vegetables"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picoutput 
      Height          =   4095
      Left            =   2760
      ScaleHeight     =   4035
      ScaleWidth      =   3795
      TabIndex        =   2
      Top             =   360
      Width           =   3855
   End
   Begin VB.CommandButton cmsshow 
      Caption         =   "Show Vegtables"
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
      Width           =   1455
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
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404080&
      Caption         =   "It is recomended that you consume 2 to 3 cups of vegetables per day"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
End
Attribute VB_Name = "frm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Healthy Living
'frm5
'Ben Morris
'March 21
'displays the different vegetables
Option Explicit
Dim veggies(1 To 20) As String
Dim CTR As Integer
Private Sub cmdback_Click()
    frm1.Show
    frm2.Hide
    frm3.Hide
    frm4.Hide
    frm5.Hide
    frm6.Hide
    frm7.Hide
    frm8.Hide
    'shows the pyramid and hides all others
End Sub

Private Sub cmsshow_Click()

    'gets vegetable examples from a file and prints them
    
    picoutput.Cls
    CTR = 0
    picoutput.Print "These are some Examples Vegetables"
    picoutput.Print "--------------------------------------------------------------------"
    Open App.Path & "\Veggies.txt" For Input As #1
    Do Until EOF(1)
    
    ' Get the data from the file
        CTR = CTR + 1
        Input #1, veggies(CTR)
        
        picoutput.Print veggies(CTR)
    
    Loop
    Close #1
End Sub

