VERSION 5.00
Begin VB.Form frmexplanation 
   BackColor       =   &H00000000&
   Caption         =   "wrestlingexplanation"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   12512.62
   ScaleMode       =   0  'User
   ScaleWidth      =   15630.77
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSortingPage 
      BackColor       =   &H000000FF&
      Caption         =   "Go to Sorting Page"
      Height          =   975
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9600
      Width           =   2415
   End
   Begin VB.CommandButton cmdNS 
      BackColor       =   &H000000FF&
      Caption         =   "Go to Searching Page"
      Height          =   975
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9600
      Width           =   2415
   End
   Begin VB.CommandButton cmdHomepage 
      BackColor       =   &H000000FF&
      Caption         =   "Go to Home Page"
      Height          =   975
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9600
      Width           =   2415
   End
   Begin VB.CommandButton cmdend 
      BackColor       =   &H000000FF&
      Caption         =   "End"
      Height          =   975
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8520
      Width           =   2415
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000007&
      FillStyle       =   0  'Solid
      Height          =   2775
      Left            =   10560
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   4395
      TabIndex        =   4
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back"
      Height          =   975
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8520
      Width           =   2415
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H000000FF&
      Caption         =   "Next"
      Height          =   975
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8520
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      Height          =   1455
      Left            =   1080
      ScaleHeight     =   1395
      ScaleWidth      =   9195
      TabIndex        =   1
      Top             =   6840
      Width           =   9255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   6255
      Left            =   1080
      ScaleHeight     =   6195
      ScaleWidth      =   8835
      TabIndex        =   0
      Top             =   360
      Width           =   8895
   End
End
Attribute VB_Name = "frmexplanation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctr As Integer

Private Sub cmdBack_Click() 'when the user clicks the back button it will go to the slide that was shown previous to it.


        ctr = ctr - 1
        If ctr <= 0 Then
            MsgBox "you must click next before back!!!", , "error"
            ctr = ctr + 1
        End If
        
         If ctr = 1 Then
            Picture2.Cls
            Picture1.Picture = LoadPicture(App.Path & "\WrestlingMat.jpg")
            Picture2.Print "The Sport of Wreslting takes place on a mat.  There are two circles which govern wrestlers while they are in competition."
            Picture2.Print "The first, and smaller circle is the place where the begining of every period is started or if the wreslters"
            Picture2.Print " leave the second larger circle then the match is restarted in the smaller of the two."
        End If
      
        If ctr = 2 Then
            Picture2.Cls
            Picture1.Picture = LoadPicture(App.Path & "\Feet.jpg")
            Picture2.Print "The start of the first of three periods starts on the feet."
            Picture2.Print "Scoring from the feet occurs when you take your opponent to the mat and are in control"
            Picture2.Print "this results in two points. "
        End If
        
        
        If ctr = 3 Then
           Picture2.Cls
            Picture1.Picture = LoadPicture(App.Path & "\Turning.jpg")
            Picture2.Print "At the start of the second and third period you have the choice of starting on top or on bottom."
            Picture2.Print "this position is called the  Referee's Position."
            Picture2.Print "From this postion you can turn your opponent for 2 or 3 points, this is called nearfall."
        End If
             
        If ctr = 4 Then
            Picture2.Cls
            Picture1.Picture = LoadPicture(App.Path & "\Pin.jpg")
            Picture2.Print " The final goal is to pin your opponent's shoulder to the mat which results in the end of the match and 6 team points"
   End If
   
   
End Sub

Private Sub cmdHomepage_Click()
frmexplanation.Visible = False 'goes from the explanation page to the homepage
frmWrestlingSorter.Visible = True
End Sub

Private Sub cmdNext_Click()
    'when the user clicks the Next button then the program will go from one slide to another explaning wreslting through pictures and captions.
  
   ctr = ctr + 1
    
         If ctr = 1 Then
            Picture1.Picture = LoadPicture(App.Path & "\WrestlingMat.jpg")
            Picture2.Print "The Sport of Wreslting takes place on a mat.  There are two circles which govern wrestlers while they are in competition."
            Picture2.Print "The first, and smaller circle is the place where the begining of every period is started or if the wreslters"
            Picture2.Print " leave the second larger circle then the match is restarted in the smaller of the two."
        End If
      
        If ctr = 2 Then
            Picture2.Cls
            Picture1.Picture = LoadPicture(App.Path & "\Feet.jpg")
            Picture2.Print "The start of the first of three periods starts on the feet."
            Picture2.Print "Scoring from the feet occurs when you take your opponent to the mat and are in control"
            Picture2.Print "this results in two points. "
        End If
        
        
        If ctr = 3 Then
           Picture2.Cls
            Picture1.Picture = LoadPicture(App.Path & "\Turning.jpg")
            Picture2.Print "At the start of the second and third period you have the choice of starting on top or on bottom."
            Picture2.Print "this position is called the  Referee's Position."
            Picture2.Print "From this postion you can turn your opponent for 2 or 3 points, this is called nearfall."
        End If
             
        If ctr = 4 Then
            Picture2.Cls
            Picture1.Picture = LoadPicture(App.Path & "\Pin.jpg")
            Picture2.Print " The final goal is to pin your opponent's shoulder to the mat which results in the end of the match and 6 team points"
     
        End If
   
   
      If ctr > 4 Then
        MsgBox "For more information reference your local Library.", , "End of Explanation"
      End If
   
End Sub

    

Private Sub cmdend_Click()
    End 'ends the program
End Sub

Private Sub cmdNS_Click()
frmexplanation.Visible = False 'goes from explanation page to the Searching page
FrmNameSearch.Visible = True
End Sub

Private Sub cmdSortingPage_Click()
frmexplanation.Visible = False 'Goes from the explanation page to the searching Page
frmSortingPage.Visible = True
End Sub
