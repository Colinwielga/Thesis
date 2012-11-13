VERSION 5.00
Begin VB.Form frmauthors 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   9750
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12090
   FillColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   1440
      Picture         =   "frm4.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   7755
      TabIndex        =   5
      Top             =   1920
      Width           =   7815
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to East High"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9720
      TabIndex        =   4
      Top             =   4200
      Width           =   2055
   End
   Begin VB.PictureBox picresults 
      Height          =   3015
      Left            =   1440
      ScaleHeight     =   2955
      ScaleWidth      =   7755
      TabIndex        =   3
      Top             =   6480
      Width           =   7815
   End
   Begin VB.CommandButton cmdmegan 
      Caption         =   "Megan"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4320
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton cmdlaura 
      Caption         =   "Laura"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdkirsten 
      Caption         =   "Kirsten"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6720
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmauthors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
'show the buttons page when the user clicks on go back to east high button
frmauthors.Hide
frmbuttons.Show
Frmcharacter.Hide
frmtitle.Hide
frmnamethattune.Hide
FrmTrivia.Hide
frmquiz.Hide
frmtitle.Hide

End Sub

Private Sub cmdkirsten_Click()
'dim variables
Dim kirsten(1 To 10) As String
Dim CTR As Integer
Dim N As Integer

'read in the text file about kirsten and print it if the user clicks on the kirsten button
CTR = 0
picresults.Cls

Open App.Path & "\kirsten.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, kirsten(CTR)
Loop
Close #1

For N = 1 To CTR
    picresults.Print kirsten(N)
Next N
End Sub

Private Sub cmdlaura_Click()
'dim variables
Dim laura(1 To 10) As String
Dim CTR As Integer
Dim N As Integer
'read in the text file about laura and print it if the user clicks on the laura button

CTR = 0
picresults.Cls

Open App.Path & "\laura.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, laura(CTR)
Loop
Close #1

For N = 1 To CTR
    picresults.Print laura(N)
Next N
End Sub

Private Sub cmdmegan_Click()
'dim variables
Dim megan(1 To 10) As String
Dim CTR As Integer
Dim N As Integer
'read in the text file about megan and print it if the user clicks on the megan button

CTR = 0
picresults.Cls

Open App.Path & "\megan.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, megan(CTR)
Loop
Close #1

For N = 1 To CTR
    picresults.Print megan(N)
Next N
End Sub
