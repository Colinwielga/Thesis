VERSION 5.00
Begin VB.Form SlideShowItaly 
   BackColor       =   &H000080FF&
   Caption         =   "SlideShowItaly"
   ClientHeight    =   10575
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   15240
   Begin VB.PictureBox picresults 
      BackColor       =   &H000080FF&
      Height          =   7815
      Left            =   1680
      ScaleHeight     =   7755
      ScaleWidth      =   12435
      TabIndex        =   3
      Top             =   2520
      Width           =   12495
   End
   Begin VB.CommandButton cmdgoback 
      Caption         =   "Look At Another City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "View Slideshow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblitaly 
      BackColor       =   &H000080FF&
      Caption         =   "What Does Italy Look Like?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "SlideShowItaly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Where to Travel in Italy
'Form Name: SlideShowItaly
'Author: Sarah Dayton
'This form is to show the user what they can see in Italy in slideshow form
Option Explicit

Private Sub cmdgoback_Click()
OpeningPage.Show
Milan.Hide
Venice.Hide
Florence.Hide
Rome.Hide
Naples.Hide
SlideShowItaly.Hide
End Sub

Private Sub cmdshow_Click()
Dim stopper As Integer, t As Double, ctr2 As Double, PicCTR As Integer, pictures(1 To 100) As String, CTR As Integer
CTR = 0
Open App.Path & "\slideshowitaly.txt" For Input As #1
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, pictures(CTR)
    Loop

PicCTR = 1

stopper = 0
Do While (stopper < 21)
    picresults.Picture = LoadPicture(App.Path & "\" & pictures(PicCTR))

    t = Timer
    Do While (Timer - t) < 2
        ctr2 = ctr2 + 1
        If ctr2 = 1000000 Then
            ctr2 = 0
        End If
    Loop

    stopper = stopper + 1
    
    PicCTR = (stopper Mod CTR) + 1
Loop
picresults.Cls
Close #1
End Sub
