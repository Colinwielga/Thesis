VERSION 5.00
Begin VB.Form frmSInfo 
   BackColor       =   &H00008080&
   Caption         =   "Second Electric Information"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Close"
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdDisplayInfo 
      Caption         =   "Show Instrument Schedule for Second Electric"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   5400
      Width           =   3855
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   240
      ScaleHeight     =   4875
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "frmSInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This displays the information inputted from the second electric form to display.

Private Sub cmdDisplayInfo_Click()
'Reads the file and displays it.
    Dim instString(1 To 100) As String
    Dim gelString(1 To 100) As String
    Dim CTR As Integer
    Dim pos As Integer
    
    picResults.Cls
    
    Open App.Path & "\Second.txt" For Input As #1
        Do Until EOF(1)
            CTR = CTR + 1
            Input #1, CTR, instString(CTR), gelString(CTR)
        Loop
        
        For pos = 1 To CTR
            picResults.Print pos, instString(pos), gelString(pos)
        Next pos
        
        Close #1

End Sub

Private Sub cmdEnd_Click()
'Closes the form.
    frmSInfo.Hide

End Sub

