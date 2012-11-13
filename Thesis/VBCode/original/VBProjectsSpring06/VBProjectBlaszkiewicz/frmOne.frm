VERSION 5.00
Begin VB.Form frmOne 
   BackColor       =   &H80000003&
   Caption         =   "Toyota Cars"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picInfo 
      Height          =   4695
      Left            =   5280
      ScaleHeight     =   4635
      ScaleWidth      =   4755
      TabIndex        =   20
      Top             =   0
      Width           =   4815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Quit"
      Height          =   1215
      Left            =   4320
      TabIndex        =   19
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Compare Cars"
      Height          =   1215
      Left            =   2040
      TabIndex        =   18
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdView6 
      Caption         =   "View Info"
      Height          =   615
      Left            =   1200
      TabIndex        =   11
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdBuy6 
      Caption         =   "Buy"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdView5 
      Caption         =   "View Info"
      Height          =   615
      Left            =   1200
      TabIndex        =   9
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdBuy5 
      Caption         =   "Buy"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdView4 
      Caption         =   "View Info"
      Height          =   615
      Left            =   1200
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cdmBuy4 
      Caption         =   "Buy"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdView3 
      Caption         =   "View Info"
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdBuy3 
      Caption         =   "Buy"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdView2 
      Caption         =   "View Info"
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cdmBuy2 
      Caption         =   "Buy"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdView1 
      Caption         =   "View Info"
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdBuy1 
      Caption         =   "Buy"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Image Image8 
      Height          =   1800
      Left            =   6960
      Picture         =   "frmOne.frx":0000
      Top             =   4680
      Width           =   3000
   End
   Begin VB.Image Image7 
      Height          =   1230
      Left            =   240
      Picture         =   "frmOne.frx":2789
      Top             =   4800
      Width           =   1485
   End
   Begin VB.Image Image6 
      Height          =   675
      Left            =   3840
      Picture         =   "frmOne.frx":31E7
      Top             =   3600
      Width           =   1650
   End
   Begin VB.Image Image5 
      Height          =   675
      Left            =   3840
      Picture         =   "frmOne.frx":6C85
      Top             =   2880
      Width           =   1650
   End
   Begin VB.Image Image4 
      Height          =   675
      Left            =   3840
      Picture         =   "frmOne.frx":A723
      Top             =   2160
      Width           =   1650
   End
   Begin VB.Image Image3 
      Height          =   675
      Left            =   3840
      Picture         =   "frmOne.frx":E1C1
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Image Image2 
      Height          =   675
      Left            =   3840
      Picture         =   "frmOne.frx":114A3
      Top             =   720
      Width           =   1650
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   3840
      Picture         =   "frmOne.frx":14F41
      Top             =   0
      Width           =   1650
   End
   Begin VB.Label lblAvalon 
      BackColor       =   &H80000003&
      Caption         =   "Toyota Avalon"
      Height          =   615
      Left            =   2280
      TabIndex        =   17
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblPrius 
      BackColor       =   &H80000003&
      Caption         =   "Toyota Prius"
      Height          =   495
      Left            =   2280
      TabIndex        =   16
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lableCamry2 
      BackColor       =   &H80000003&
      Caption         =   "Toyota Camry Solara"
      Height          =   495
      Left            =   2280
      TabIndex        =   15
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblCamry 
      BackColor       =   &H80000003&
      Caption         =   "Toyota Camry"
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblMatrix 
      BackColor       =   &H80000003&
      Caption         =   "Toyota Matrix"
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblCorrola 
      BackColor       =   &H80000003&
      Caption         =   "Toyota Corolla"
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Data(1 To 100) As String

Private Sub cdmBuy2_Click()
    frmOne.Visible = False
    frmThree.Visible = True
End Sub

Private Sub cdmBuy4_Click()
    frmOne.Visible = False
    frmThree.Visible = True
End Sub

Private Sub cmdBuy1_Click()
    frmOne.Visible = False
    frmThree.Visible = True
End Sub

Private Sub cmdBuy3_Click()
    frmOne.Visible = False
    frmThree.Visible = True
End Sub

Private Sub cmdBuy5_Click()
    frmOne.Visible = False
    frmThree.Visible = True
End Sub

Private Sub cmdBuy6_Click()
    frmOne.Visible = False
    frmThree.Visible = True
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdMove_Click()
    frmOne.Visible = False
    frmTwo.Visible = True
End Sub

Private Sub cmdView1_Click()
    Dim Pos, Size As Integer
    Pos = 0
    Open App.Path & "\Corolla.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Data(Pos)
    Loop
    Close #1
    Size = Pos
    picInfo.Cls
    For Pos = 1 To Size
        picInfo.Print Data(Pos)
    Next Pos
    
End Sub

Private Sub cmdView2_Click()
    Dim Pos, Size As Integer
    Pos = 0
    Open App.Path & "\Matrix.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Data(Pos)
    Loop
    Close #1
    Size = Pos
    picInfo.Cls
    For Pos = 1 To Size
        picInfo.Print Data(Pos)
    Next Pos
End Sub

Private Sub cmdView3_Click()
    Dim Pos, Size As Integer
    Pos = 0
    Open App.Path & "\Camry.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Data(Pos)
    Loop
    Close #1
    Size = Pos
    picInfo.Cls
    For Pos = 1 To Size
        picInfo.Print Data(Pos)
    Next Pos
End Sub

Private Sub cmdView4_Click()
    Dim Pos, Size As Integer
    Pos = 0
    Open App.Path & "\CamrySolara.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Data(Pos)
    Loop
    Close #1
    Size = Pos
    picInfo.Cls
    For Pos = 1 To Size
        picInfo.Print Data(Pos)
    Next Pos
End Sub

Private Sub cmdView5_Click()
    Dim Pos, Size As Integer
    Pos = 0
    Open App.Path & "\Prius.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Data(Pos)
    Loop
    Close #1
    Size = Pos
    picInfo.Cls
    For Pos = 1 To Size
        picInfo.Print Data(Pos)
    Next Pos
End Sub

Private Sub cmdView6_Click()
    Dim Pos, Size As Integer
    Pos = 0
    Open App.Path & "\Avalon.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Data(Pos)
    Loop
    Close #1
    Size = Pos
    picInfo.Cls
    For Pos = 1 To Size
        picInfo.Print Data(Pos)
    Next Pos
End Sub

Private Sub Form_Load()
    MsgBox "WELCOME TO TOYOTA STORE!!! (Click OK to begin)", , "Toyota Cars"
End Sub
