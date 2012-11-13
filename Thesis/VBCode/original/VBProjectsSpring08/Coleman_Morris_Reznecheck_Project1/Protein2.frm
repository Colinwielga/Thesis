VERSION 5.00
Begin VB.Form frm7 
   BackColor       =   &H000080FF&
   Caption         =   "Protein"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picoutput 
      Height          =   4455
      Left            =   2640
      ScaleHeight     =   4395
      ScaleWidth      =   3915
      TabIndex        =   2
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Protein"
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
      Left            =   720
      TabIndex        =   1
      Top             =   1560
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
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "It is recomended that you have 5 to 6 ounces of protein each day"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
End
Attribute VB_Name = "frm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Healthy Living
'frm7
'Ben Morris
'March 21
'displays the different vegetables
Dim protein(1 To 10) As String
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
    picoutput.Print "These are some Examples of Protein"
    picoutput.Print "--------------------------------------------------------------------"
    Open App.Path & "\Protein.txt" For Input As #1
    Do Until EOF(1)
    
    ' Get the data from the file
        CTR = CTR + 1
        Input #1, protein(CTR)
        
        picoutput.Print protein(CTR)
    
    Loop
    Close #1
End Sub
