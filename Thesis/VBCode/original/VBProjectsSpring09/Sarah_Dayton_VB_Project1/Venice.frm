VERSION 5.00
Begin VB.Form Venice 
   BackColor       =   &H00C0C000&
   Caption         =   "Form3"
   ClientHeight    =   10575
   ClientLeft      =   1950
   ClientTop       =   540
   ClientWidth     =   11880
   LinkTopic       =   "Form3"
   ScaleHeight     =   10575
   ScaleWidth      =   11880
   Begin VB.PictureBox picresults 
      BackColor       =   &H00C0C000&
      Height          =   6255
      Left            =   1200
      ScaleHeight     =   6195
      ScaleWidth      =   9795
      TabIndex        =   3
      Top             =   2400
      Width           =   9855
   End
   Begin VB.CommandButton cmdgoback 
      Caption         =   "Look For Another City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdslide 
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
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "What Do To Do In Venice?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   855
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Venice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Where to Travel in Italy
'Form Name: Venice
'Author: Sarah Dayton
'This form shows a slideshow of sights for Venice
Option Explicit

Private Sub cmdgoback_Click()
Close #1
OpeningPage.Show
Milan.Hide
Venice.Hide
Florence.Hide
Rome.Hide
Naples.Hide
SlideShowItaly.Hide
End Sub

Private Sub cmdslide_Click()
Dim stopper As Integer, t As Double, ctr2 As Double, PicCTR As Integer, pictures(1 To 100) As String, CTR As Integer
CTR = 0
Open App.Path & "\Venice.txt" For Input As #1
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, pictures(CTR)
    Loop

PicCTR = 1

stopper = 0
Do While (stopper < 7)
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
End Sub
