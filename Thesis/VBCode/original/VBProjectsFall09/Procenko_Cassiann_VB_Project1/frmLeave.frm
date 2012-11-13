VERSION 5.00
Begin VB.Form frmLeave 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Works Cited"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWorkCited 
      Caption         =   "View Works Cited"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtYesNo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Text            =   "Yes or No"
      Top             =   960
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   3240
      ScaleHeight     =   4755
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   600
      Width           =   6255
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Leave/Quit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label lblInstruct 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Did you enjoy your visit? (Enter either Yes or No)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Music VB Project by Cassiann Procenko
'Form Name is frmLeave
'Date written 10/16/2009
'Purpose of this form is to show the viewer the works cited information and where I got all the information about the animals.


Private Sub cmdQuit_Click()
'end program
End
End Sub

Private Sub cmdSubmit_Click()
'define variables
Dim Answer As String

'if then for textbox answer
Answer = txtYesNo
    If Answer = "Yes" Then
        MsgBox "That is great. Please visit again soon!"
    ElseIf Answer = "No" Then
        MsgBox "I am sorry you did not enjoy your visit."
    Else
        MsgBox "You did not enter a valid answer."
    End If
    
End Sub

Private Sub cmdWorkCited_Click()
'defining variables
Dim websiteList(1 To 50) As String
Dim CTR As Integer

'print header information
picResults.Print
picResults.Print "Websites Used"
picResults.Print "***************************"

'open the file to be read and made into array
Open App.Path & "\WorksCited.txt" For Input As #1

CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, websiteList(CTR)
    picResults.Print websiteList(CTR)
Loop

Close #1
End Sub
