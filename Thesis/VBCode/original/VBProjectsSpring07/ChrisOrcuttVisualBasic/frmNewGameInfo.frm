VERSION 5.00
Begin VB.Form frmGameReviews 
   Caption         =   "Latest Game Reviews"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   13620
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   6735
      Left            =   120
      ScaleHeight     =   6675
      ScaleWidth      =   13155
      TabIndex        =   1
      Top             =   120
      Width           =   13215
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   6960
      Width           =   1935
   End
End
Attribute VB_Name = "frmGameReviews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdReturn_Click()
    frmGameReviews.Hide
    frmSelectWant.Show
End Sub
Private Sub Form_Load()
     Dim Ctr As Integer
        Open App.Path & "\CrackdownReview.txt" For Input As #1
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, CrackdownReview(Ctr)
            picResults.Print ; CrackdownReview(Ctr)
            Loop
        Close #1
End Sub
