VERSION 5.00
Begin VB.Form Frmcharacter 
   Caption         =   "Form2"
   ClientHeight    =   9240
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   Picture         =   "Frm2.frx":0000
   ScaleHeight     =   9240
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
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
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   7440
      Width           =   1455
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H80000009&
      Height          =   3735
      Left            =   3120
      ScaleHeight     =   3675
      ScaleWidth      =   5235
      TabIndex        =   4
      Top             =   3720
      Width           =   5295
   End
   Begin VB.CommandButton cmdfindout 
      Caption         =   "Click to Find out More!"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   10560
      TabIndex        =   3
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox txtcharacter 
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   8640
      Width           =   3975
   End
   Begin VB.Label lblnames 
      BackColor       =   &H80000009&
      Caption         =   $"Frm2.frx":161ECE
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblcharacter 
      BackColor       =   &H80000009&
      Caption         =   "Enter your Favorite Character's Name to Find out More:"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   8760
      Width           =   4455
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H80000009&
      Caption         =   "Meet the East High Gang!"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Frmcharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdback_Click()
frmauthors.Hide
frmbuttons.Hide
Frmcharacter.Hide
frmtitle.Hide
frmnamethattune.Hide
FrmTrivia.Hide
frmquiz.Hide
frmtitle.Show
End Sub

Private Sub cmdfindout_Click()
Dim name(1 To 100) As String
Dim text As String
Dim ctr As Integer
Dim Found As Boolean
Dim CName As String
Dim pos As Integer
Dim Info(1 To 100) As String
Dim NumLines As Integer
Dim NewLine As String

Open App.Path & "\hsm.txt" For Input As #1
    ctr = 0
    CName = txtcharacter.text
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, name(ctr), NumLines

        For pos = 1 To NumLines
            Input #1, NewLine
            Info(ctr) = Info(ctr) & vbCrLf & NewLine
        Next pos
    Loop
    Close #1
    pos = 0
    Do While (Found = False And pos < ctr)
        pos = pos + 1
        If LCase(name(pos)) = LCase(CName) Then
            Found = True
        End If
    Loop
    picresults.Cls
    If Found = True Then
        picresults.Print Info(pos)
    Else
        MsgBox "Please try again.  Make sure name is properly spelled.", , "Error"
    End If
    End Sub
