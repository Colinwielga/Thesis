VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11295
   FillColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit Looking for New Cars"
      Height          =   1695
      Left            =   4200
      TabIndex        =   8
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton cmdsedan 
      Caption         =   "Sedan"
      Height          =   975
      Left            =   7200
      TabIndex        =   7
      Top             =   7200
      Width           =   3135
   End
   Begin VB.CommandButton cmdTruck 
      Caption         =   "Truck"
      Height          =   975
      Left            =   360
      TabIndex        =   6
      Top             =   7200
      Width           =   2535
   End
   Begin VB.CommandButton cmdSUV 
      Caption         =   "Sport Utility "
      Height          =   975
      Left            =   7440
      TabIndex        =   5
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton cmdsport 
      Caption         =   "Sports Cars"
      Height          =   975
      Left            =   480
      TabIndex        =   4
      Top             =   2880
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   240
      Picture         =   "VB_Opening_page.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   3555
      TabIndex        =   3
      Top             =   240
      Width           =   3615
   End
   Begin VB.PictureBox Picture4 
      Height          =   2295
      Left            =   6960
      Picture         =   "VB_Opening_page.frx":1A8F2
      ScaleHeight     =   2235
      ScaleWidth      =   3555
      TabIndex        =   2
      Top             =   4800
      Width           =   3615
   End
   Begin VB.PictureBox Picture3 
      Height          =   2295
      Left            =   240
      Picture         =   "VB_Opening_page.frx":351E4
      ScaleHeight     =   2235
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   4800
      Width           =   3615
   End
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   6960
      Picture         =   "VB_Opening_page.frx":4FAD6
      ScaleHeight     =   2235
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   $"VB_Opening_page.frx":6F558
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3240
      TabIndex        =   9
      Top             =   2760
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(PROJECT:OPENING PAGE)
'(FORM:NARROWING DOWN THE TYPE OF AUTOMOBILE)
'TORY BERTELSON
'10-23-05
'OBJECTIVE: THIS FORM ALLOWS THE USER TO CHOOSE A VARIOUS TYPE OF AUTOMOBILES HE IS INTERESTED IN


Dim price As Integer
Dim CTR As Integer
Dim modelSport(1 To 100) As String
Dim priceSport(1 To 100) As Single
Dim gasSport(1 To 100) As Integer


Private Sub cmdquit_Click()

MsgBox "you will now leave the car search", , "done looking" 'lets the user know he is quiting the car search
End
End Sub

Private Sub cmdsedan_Click()    'allows the user to enter the Sedan page
sedan.Visible = True
Sport.Visible = False
SUV.Visible = False
Truck.Visible = False

End Sub

Private Sub cmdsport_Click()    'allows the user to enter the Sports car page
Sport.Visible = True
sedan.Visible = False
SUV.Visible = False
Truck.Visible = False


        
        
End Sub

Private Sub cmdSUV_Click()  'allows the user to enter the SUV page
SUV.Visible = True
Sport.Visible = False
sedan.Visible = False
Truck.Visible = False


End Sub

Private Sub cmdTruck_Click()    'allows teh user to enter the Truck page
Truck.Visible = True
Sport.Visible = False
sedan.Visible = False
SUV.Visible = False

End Sub
