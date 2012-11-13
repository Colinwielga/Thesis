VERSION 5.00
Begin VB.Form frm6 
   BackColor       =   &H00008000&
   Caption         =   "Dairy"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picoutput 
      Height          =   4575
      Left            =   2400
      ScaleHeight     =   4515
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Dairy"
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
      TabIndex        =   1
      Top             =   1320
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
      BackColor       =   &H00008000&
      Caption         =   "It is recomended that you have three cups of dairy in your diet each day"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
End
Attribute VB_Name = "frm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Healthy Living
'frm6
'Ben Morris
'March 21
'displays the different dairy
Dim dairy(1 To 6) As String
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

Private Sub Command1_Click()
    picoutput.Cls
    CTR = 0
    picoutput.Print "These are some Examples of Dairy"
    picoutput.Print "--------------------------------------------------------------------"
    Open App.Path & "\Dairy.txt" For Input As #1
    Do Until EOF(1)
    
    ' Get the data from the file
        CTR = CTR + 1
        Input #1, dairy(CTR)
        
        picoutput.Print dairy(CTR)
    
    Loop
    Close #1
End Sub
