VERSION 5.00
Begin VB.Form frmBigService 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000C000&
      Caption         =   "Quit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   4455
   End
   Begin VB.CommandButton cmdGotoSocial 
      BackColor       =   &H0000C000&
      Caption         =   "Go calculate Social Points"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   4455
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H0000C000&
      Caption         =   "list events"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   4455
   End
   Begin VB.CommandButton cmdService 
      BackColor       =   &H0000C000&
      Caption         =   "calcuate your service points"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   4455
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C0FF&
      Height          =   4935
      Left            =   6120
      ScaleHeight     =   4875
      ScaleWidth      =   3195
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lblBigService 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Service Points!"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   7695
   End
End
Attribute VB_Name = "frmBigService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim Service(1 To 25) As String
 Dim ctr As Integer

Private Sub cmdGotoSocial_Click()
    frmBigSocial.Show
    frmBigService.Hide
End Sub

Private Sub cmdList_Click()
    Open App.Path & "\service.txt" For Input As #1
    ctr = 0
    Do While Not EOF(1)
        ctr = ctr + 1
        Input #1, Service(ctr)
        picResults.Print Service(ctr)
    Loop
    picResults.Print
    picResults.Print
    cmdList.Enabled = False
    Close #1
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdService_Click()
Dim points As Integer
    points = InputBox("Enter the number of service events you have attended", "Events attened")
    If points > 11 Then
        MsgBox "That is an invailed amount of events", , "Error"
    Else
        Runningtotal = points * 10
        BServiceCTR = Runningtotal
        picResults.Print "You have"; BServiceCTR; "service points."
    End If
    
    cmdGotoSocial.Enabled = True
    cmdQuit.Enabled = True
End Sub
