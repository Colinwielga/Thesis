VERSION 5.00
Begin VB.Form frmIdentification 
   BackColor       =   &H000000C0&
   Caption         =   "Indentifying special limit ducks"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   13005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReference 
      Caption         =   "Load References"
      Height          =   615
      Left            =   360
      TabIndex        =   14
      Top             =   1320
      Width           =   3495
   End
   Begin VB.PictureBox picReference 
      Height          =   735
      Left            =   360
      ScaleHeight     =   675
      ScaleWidth      =   5115
      TabIndex        =   12
      Top             =   7080
      Width           =   5175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit Program"
      Height          =   615
      Left            =   360
      TabIndex        =   11
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdStates 
      Caption         =   "Pick a State"
      Height          =   615
      Left            =   2280
      TabIndex        =   10
      Top             =   5400
      Width           =   1575
   End
   Begin VB.PictureBox picDuckPicture 
      Height          =   7575
      Left            =   6000
      ScaleHeight     =   7515
      ScaleWidth      =   6315
      TabIndex        =   9
      Top             =   720
      Width           =   6375
   End
   Begin VB.CommandButton cmdPintail 
      Caption         =   "Pintails"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdBlackDuck 
      Caption         =   "Black Ducks"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2280
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdCanvasback 
      Caption         =   "Canvasbacks"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdScaup 
      Caption         =   "Scaup"
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdRedhead 
      Caption         =   "Redheads"
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdWoodDuck 
      Caption         =   "Wood Ducks"
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdMallard 
      Caption         =   "Mallards"
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblReference 
      BackStyle       =   0  'Transparent
      Caption         =   "Reference for picture."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   13
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "Note:  Ducks listed are special limit                   ducks in at least one state."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label lblIdentification 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pick a duck that you want to identify and its picture will appear."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "frmIdentification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pictureReference(1 To 10) As String, DuckName(1 To 10) As String
Dim Ctr As Integer
Dim Found As Boolean


Private Sub cmdBlackDuck_Click()
'Loads picture of the duck into a picture box and the website it was taken off of into another picture box

picDuckPicture.Picture = LoadPicture(App.Path & "\blackduck1.gif")
Dim Ctr2 As Integer
Ctr2 = 0

Found = False

'finds the correct Duck

Do Until Found = True
    Ctr2 = Ctr2 + 1
    If DuckName(Ctr2) = "BlackDuck" Then
        Found = True
    End If
Loop
picReference.Cls
picReference.Print pictureReference(Ctr2)

End Sub

Private Sub cmdCanvasback_Click()
'Loads picture of the duck into a picture box and the website it was taken off of into another picture box
picDuckPicture.Picture = LoadPicture(App.Path & "\canvasback1.gif")
Dim Ctr2 As Integer
Ctr2 = 0

Found = False

'finds the correct Duck

Do Until Found = True
    Ctr2 = Ctr2 + 1
    If DuckName(Ctr2) = "Canvasback" Then
        Found = True
    End If
Loop
picReference.Cls
picReference.Print pictureReference(Ctr2)

End Sub

Private Sub cmdMallard_Click()
'Loads picture of the duck into a picture box and the website it was taken off of into another picture box
picDuckPicture.Picture = LoadPicture(App.Path & "\mallard1.gif")
Dim Ctr2 As Integer
Ctr2 = 0

Found = False

'finds the correct Duck

Do Until Found = True
    Ctr2 = Ctr2 + 1
    If DuckName(Ctr2) = "Mallard" Then
        Found = True
    End If
Loop
picReference.Cls
picReference.Print pictureReference(Ctr2)

End Sub

Private Sub cmdPintail_Click()
'Loads picture of the duck into a picture box and the website it was taken off of into another picture box
picDuckPicture.Picture = LoadPicture(App.Path & "\pintail1.gif")
Dim Ctr2 As Integer
Ctr2 = 0

Found = False

'finds the correct Duck

Do Until Found = True
    Ctr2 = Ctr2 + 1
    If DuckName(Ctr2) = "Pintail" Then
        Found = True
    End If
Loop
picReference.Cls
picReference.Print pictureReference(Ctr2)

End Sub

Private Sub cmdQuit_Click()
'ends the program
End
End Sub

Private Sub cmdRedhead_Click()
'Loads picture of the duck into a picture box and the website it was taken off of into another picture box
picDuckPicture.Picture = LoadPicture(App.Path & "\redhead1.gif")
Dim Ctr2 As Integer
Ctr2 = 0

Found = False

'finds the correct Duck

Do Until Found = True
    Ctr2 = Ctr2 + 1
    If DuckName(Ctr2) = "Redhead" Then
        Found = True
    End If
Loop
picReference.Cls
picReference.Print pictureReference(Ctr2)

End Sub

Private Sub cmdReference_Click()
'opens a file with the names of the special limits ducks and the reference for the picture
 
Open App.Path & "/pictureReference.txt" For Input As #1

Ctr = 0

Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, DuckName(Ctr), pictureReference(Ctr)
Loop
    
Close #1

cmdReference.Enabled = False
cmdScaup.Enabled = True
cmdWoodDuck.Enabled = True
cmdRedhead.Enabled = True
cmdPintail.Enabled = True
cmdMallard.Enabled = True
cmdCanvasback.Enabled = True
cmdBlackDuck.Enabled = True

End Sub

Private Sub cmdStates_Click()
'takes user to the "states" page
frmIdentification.Hide
frmBegining.Show
End Sub

Private Sub cmdScaup_Click()
'Loads picture of the duck into a picture box and the website it was taken off of into another picture box
picDuckPicture.Picture = LoadPicture(App.Path & "\scaup1.gif")
Dim Ctr2 As Integer
Ctr2 = 0

Found = False

'finds the correct Duck

Do Until Found = True
    Ctr2 = Ctr2 + 1
    If DuckName(Ctr2) = "Scaup" Then
        Found = True
    End If
Loop
picReference.Cls
picReference.Print pictureReference(Ctr2)

End Sub

Private Sub cmdWoodDuck_Click()
'Loads picture of the duck into a picture box and the website it was taken off of into another picture box
picDuckPicture.Picture = LoadPicture(App.Path & "\woodduck1.gif")
Dim Ctr2 As Integer
Ctr2 = 0

Found = False

'finds the correct Duck

Do Until Found = True
    Ctr2 = Ctr2 + 1
    If DuckName(Ctr2) = "WoodDuck" Then
        Found = True
    End If
Loop
picReference.Cls
picReference.Print pictureReference(Ctr2)

End Sub

