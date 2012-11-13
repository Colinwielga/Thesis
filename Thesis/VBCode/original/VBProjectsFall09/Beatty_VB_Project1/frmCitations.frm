VERSION 5.00
Begin VB.Form frmCitations 
   BackColor       =   &H80000011&
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click for Citations"
      Height          =   975
      Left            =   4680
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   3495
      Left            =   360
      ScaleHeight     =   3435
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   5160
      Width           =   975
   End
End
Attribute VB_Name = "frmCitations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: quick Facts about the MIAC'
    'Form name:Citations
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give information about where i got my information for this project.'
Private Sub cmdClick_Click()
Dim Info As String
    
    Open App.Path & "\citations.txt" For Input As #10
    
    Do While Not EOF(10)
        Input #10, Info
        picResults.Print Info
    Loop
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmCitations.Hide
    frmMIAC.Show
End Sub

