VERSION 5.00
Begin VB.Form frmworks 
   BackColor       =   &H000000FF&
   Caption         =   "Works Cited"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdalpha 
      Caption         =   "Alphabatize!"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to East High"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   3600
      Width           =   1935
   End
   Begin VB.PictureBox picresults 
      Height          =   6015
      Left            =   2760
      ScaleHeight     =   5955
      ScaleWidth      =   5955
      TabIndex        =   1
      Top             =   360
      Width           =   6015
   End
   Begin VB.CommandButton cmdworks 
      Caption         =   "Works Cited"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   2400
      Width           =   1935
   End
End
Attribute VB_Name = "frmworks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim works(1 To 40) As String
Dim N As Integer
Dim CTR As Integer


Private Sub cmdalpha_Click()
'dim variables
Dim Tempworks As String
Dim Pass As Integer
Dim Pos As Integer

'alphabatize all websites
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If works(Pos) > works(Pos + 1) Then
            Tempworks = works(Pos)
            works(Pos) = works(Pos + 1)
            works(Pos + 1) = Tempworks
        End If
    Next Pos
Next Pass
picresults.Cls
For N = 1 To CTR
picresults.Print works(N)
Next N



End Sub

Private Sub cmdback_Click()
frmauthors.Hide
frmbuttons.Hide
Frmcharacter.Hide
frmtitle.Show
frmnamethattune.Hide
FrmTrivia.Hide
frmquiz.Hide
frmworks.Hide
'show the title page when user clicks on the back to east high button

End Sub

Private Sub cmdworks_Click()

CTR = 0
picresults.Cls
'read in and open the works cited text file. Print them as well
Open App.Path & "\workscited.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, works(CTR)
Loop
Close #1

For N = 1 To CTR
    picresults.Print works(N)
Next N
cmdalpha.Enabled = True

End Sub
