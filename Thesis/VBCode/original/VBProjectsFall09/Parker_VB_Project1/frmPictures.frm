VERSION 5.00
Begin VB.Form frmPictures 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pictures"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   12270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   615
      Left            =   7800
      TabIndex        =   9
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   5400
      TabIndex        =   8
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdPic6 
      Caption         =   "Subcompact"
      Height          =   975
      Left            =   11040
      TabIndex        =   6
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdPic5 
      Caption         =   "Compact"
      Height          =   975
      Left            =   11040
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdPic4 
      Caption         =   "Coupe"
      Height          =   975
      Left            =   11040
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdPic3 
      BackColor       =   &H00000080&
      Caption         =   "Truck"
      Height          =   975
      Left            =   480
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdPic2 
      Caption         =   "Sedan"
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdPic1 
      Caption         =   "SUV"
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      Height          =   4695
      Left            =   2160
      ScaleHeight     =   4635
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   600
      Width           =   8055
   End
End
Attribute VB_Name = "frmPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name is WickedFunCarBuilder
'Form Name is frmPictures
'Author is Dan Parker
'Date written 10/18/09
'The purpose of this form is to show the user via pictures what body styles they will be able
'build upon in the next section of the program.

Private Sub cmdPic1_Click()
    picResults.Picture = LoadPicture(App.Path & "\SUV.jpg") 'loads picture of a SUV into picture space
End Sub

Private Sub cmdpic2_Click()
    picResults.Picture = LoadPicture(App.Path & "\sedan.jpg") 'loads picture of a sedan into picture space
End Sub

Private Sub cmdpic3_Click()
    picResults.Picture = LoadPicture(App.Path & "\truck.jpg") 'loads picture of a truck into picture space
End Sub

Private Sub cmdpic4_Click()
    picResults.Picture = LoadPicture(App.Path & "\coupe.jpg") 'loads picture of a coupe into picture space
End Sub

Private Sub cmdpic5_Click()
    picResults.Picture = LoadPicture(App.Path & "\compact.jpg") 'loads picture of a compact into picture space
End Sub

Private Sub cmdPic6_Click()
    picResults.Picture = LoadPicture(App.Path & "\subcompact.jpg") 'loads picture of a subcompact into picture space
End Sub
Private Sub cmdQuit_Click()
MsgBox ("Thanks for using the Wicked Fun Car Builder, " & " " & UserName & "!")
End 'quits program
End Sub

Private Sub cmdBack_Click()
    frmPictures.Hide 'hides the pictures page from user
    frmFirst.Show 'shows the first page to user
End Sub

Private Sub cmdContinue_Click()
    frmPictures.Hide 'hides the pictures page from user
    frmBegin.Show 'takes the user to the next page
End Sub


Private Sub Form_Activate()
    picResults.Picture = LoadPicture("") 'clears the picture space each time the page is opened/activated
End Sub
