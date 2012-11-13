VERSION 5.00
Begin VB.Form frmArt 
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picArtwork 
      Height          =   5895
      Left            =   480
      ScaleHeight     =   5835
      ScaleWidth      =   7995
      TabIndex        =   1
      Top             =   1320
      Width           =   8055
   End
   Begin VB.CommandButton cmdSlideshow 
      Caption         =   "slideshow"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSlideshow_Click()

Dim whichOne As Integer, stopper As Integer, t As Double, oldOne As Integer, ctr2 As Double


whichOne = 1

stopper = 0

Do While (stopper < 26)
    
    picArtwork.Picture = LoadPicture(App.Path & "\" & Art(whichOne))
    
    
    t = Timer
    Do While (Timer - t) < 2
        ctr2 = ctr2 + 1
        If ctr2 = 1000000 Then
            
            ctr2 = 0
        End If
    Loop
    
   
    stopper = stopper + 1
    
    oldOne = whichOne
   
    whichOne = (stopper Mod ctr) + 1
    
Loop

lblFileName.Caption = Art(oldOne)
lblFileName.Visible = True
 
    


End Sub
