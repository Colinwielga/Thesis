VERSION 5.00
Begin VB.Form frmFixtures 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAlph 
      BackColor       =   &H00FF00FF&
      Caption         =   "Sort by Length of Name"
      Height          =   1215
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FF00FF&
      Caption         =   "View"
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8160
      Width           =   1095
   End
   Begin VB.TextBox txtPicNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   9
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdWatts 
      BackColor       =   &H00FF00FF&
      Caption         =   "Sort by Wattage"
      Height          =   1215
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00FF00FF&
      Caption         =   "Print a list of lighting fixtures"
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   5415
      Left            =   2880
      ScaleHeight     =   5355
      ScaleWidth      =   7395
      TabIndex        =   1
      Top             =   1680
      Width           =   7455
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main Menu"
      Height          =   975
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      Width           =   3495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter the number corresponding to the light from the list about that you wish to see."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   7320
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      Caption         =   "5) Source 4 Par"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      Caption         =   "4) Par Can"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "3) Cyc Light"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "2) Fresnel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "1) Source 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   2295
   End
End
Attribute VB_Name = "frmFixtures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Theater Lighting
'Form Name: frmFixtures
'Author: Kurt Oostra
'Date Written:3/11/08
'Objective: Learn about common theater lighting fixtures
Option Explicit
Dim Pos As Integer, j As Integer, Pass As Integer, TempName As String, TempWatts As Single, TempUse As String

Private Sub cmdAlph_Click()
Picture1.Cls
Picture1.Print "Fixture"; Tab(20); "Max Wattage"; Tab(40); "Spot or Flood"
Picture1.Print "**********************************************************"
For Pass = 1 To ctr - 1
    For Pos = 1 To ctr - Pass
        If Len(names(Pos)) < Len(names(Pos + 1)) Then
            TempWatts = watts(Pos)
            watts(Pos) = watts(Pos + 1)
            watts(Pos + 1) = TempWatts
            TempName = names(Pos)
            names(Pos) = names(Pos + 1)
            names(Pos + 1) = TempName
            TempUse = use(Pos)
            use(Pos) = use(Pos + 1)
            use(Pos + 1) = TempUse
        End If
    Next
Next
For j = 1 To ctr
    Picture1.Print names(j); Tab(20); watts(j); Tab(40); use(j)
Next
End Sub

Private Sub cmdLoad_Click()
Dim j As Integer
Picture1.Cls
'prints the information on the lights, with a heading
Picture1.Print "Fixture"; Tab(20); "Max Wattage"; Tab(40); "Spot or Flood"
Picture1.Print "**********************************************************"
For j = 1 To 7
    Picture1.Print names(j); Tab(20); watts(j); Tab(40); use(j)
Next j
End Sub

Private Sub cmdReturn_Click()
'Returns to main menu
frmMainMenu.Show
frmFixtures.Hide
End Sub

Private Sub cmdView_Click()
Dim Number As Single
Number = txtPicNumber
Picture1.Cls
'Prints the picture based on the number you type in.
Select Case Number
    Case Is = 1
    Picture1.Picture = LoadPicture(App.Path & "\source4.jpg")
    Case Is = 2
    Picture1.Picture = LoadPicture(App.Path & "\Fresnelite.jpg")
    Case Is = 3
    Picture1.Picture = LoadPicture(App.Path & "\Cyc.jpg")
    Case Is = 4
    Picture1.Picture = LoadPicture(App.Path & "\ParCan.jpg")
    Case Is = 5
    Picture1.Picture = LoadPicture(App.Path & "\Source4Par.jpg")
End Select
End Sub

Private Sub cmdWatts_Click()
Picture1.Cls
'sorts the array by highest wattage
Picture1.Print "Fixture"; Tab(20); "Max Wattage"; Tab(40); "Spot or Flood"
Picture1.Print "**********************************************************"
For Pass = 1 To ctr - 1
    For Pos = 1 To ctr - Pass
        If watts(Pos) > watts(Pos + 1) Then
            TempWatts = watts(Pos)
            watts(Pos) = watts(Pos + 1)
            watts(Pos + 1) = TempWatts
            TempName = names(Pos)
            names(Pos) = names(Pos + 1)
            names(Pos + 1) = TempName
            TempUse = use(Pos)
            use(Pos) = use(Pos + 1)
            use(Pos + 1) = TempUse
        End If
    Next
Next
For j = 1 To ctr
    Picture1.Print names(j); Tab(20); watts(j); Tab(40); use(j)
Next
End Sub

Private Sub Command1_Click()

End Sub
