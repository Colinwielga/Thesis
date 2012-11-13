VERSION 5.00
Begin VB.Form frmData 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   13200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14625
   LinkTopic       =   "Form1"
   ScaleHeight     =   13200
   ScaleWidth      =   14625
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   6615
      Left            =   3120
      Picture         =   "Data.frx":0000
      ScaleHeight     =   6555
      ScaleWidth      =   7755
      TabIndex        =   8
      Top             =   6120
      Width           =   7815
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go to Main Menu"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton cmdSort3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort by Weight in Pounds"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   3135
   End
   Begin VB.CommandButton cmdSort2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort by Height in Inches"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   3135
   End
   Begin VB.CommandButton cmdSort1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort Alphabetically by First Name"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   3135
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go to Next Form"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton cmdGet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Packer's Receiver's Data"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   3720
      ScaleHeight     =   3075
      ScaleWidth      =   10155
      TabIndex        =   0
      Top             =   120
      Width           =   10215
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   2  'Horizontal Line
      Height          =   2295
      Left            =   3960
      Top             =   3360
      Width           =   9735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   2  'Horizontal Line
      Height          =   5415
      Left            =   120
      Top             =   0
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   3  'Vertical Line
      Height          =   7935
      Left            =   120
      Top             =   5760
      Width           =   13695
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Get to know the Packers' Receivers
'frmData
'Sam Pederson
'2/17/10
'This form enters the receiver's personal information from a file called packdata.txt


Option Explicit
Dim Pos As Integer, X As Integer, Pass As Integer, Temp As String

Private Sub cmdGet_Click() 'this button enters the data from the file
    Dim T As String
    picResults.Cls
    picResults.Print "Name"; Tab(20); "Jersey Number"; Tab(40); "Years of Experience in the NFL"; Tab(75); "Height (in Inches)"; Tab(95); "Weight (in Pounds)"
    picResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    Open App.Path & "\packdata.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Names(CTR), Jersey(CTR), Experience(CTR), Tall(CTR), Weight(CTR)
        picResults.Print Names(CTR); Tab(20); Jersey(CTR); Tab(40); Experience(CTR); Tab(75); Tall(CTR); Tab(95); Weight(CTR)
    Loop
    Close #1
    T = MsgBox("The data has been entered. Thank you. Enjoy!", , "What is Up")
    cmdGet.Enabled = False
    cmdSort1.Enabled = True
    cmdSort2.Enabled = True
    cmdSort3.Enabled = True
End Sub

Private Sub cmdMain_Click() 'this button takes you to the main menu
    frmWelcome.Hide
    frmMenu.Show
    frmPoll.Hide
    frmData.Hide
    frmSwap.Hide
    frmPics.Hide
    frmMusic.Hide
    frmLast.Hide
End Sub

Private Sub cmdQuit_Click() 'this button ends the program
    End
End Sub

Private Sub cmdShow_Click() 'this button takes you to the next form
    frmWelcome.Hide
    frmMenu.Hide
    frmPoll.Hide
    frmData.Hide
    frmSwap.Show
    frmPics.Hide
    frmMusic.Hide
    frmLast.Hide
End Sub

Private Sub cmdSort1_Click() 'this button sorts the names alphabetically
    picResults.Cls
    
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Names(Pos) > Names(Pos + 1) Then
                Temp = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = Temp
                Temp = Jersey(Pos)
                Jersey(Pos) = Jersey(Pos + 1)
                Jersey(Pos + 1) = Temp
                Temp = Experience(Pos)
                Experience(Pos) = Experience(Pos + 1)
                Experience(Pos + 1) = Temp
                Temp = Tall(Pos)
                Tall(Pos) = Tall(Pos + 1)
                Tall(Pos + 1) = Temp
                Temp = Weight(Pos)
                Weight(Pos) = Weight(Pos + 1)
                Weight(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    picResults.Print "Name"; Tab(20); "Jersey Number"; Tab(40); "Years of Experience in the NFL"; Tab(75); "Height (in Inches)"; Tab(95); "Weight (in Pounds)"
    picResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    For X = 1 To 7
        picResults.Print Names(X); Tab(20); Jersey(X); Tab(40); Experience(X); Tab(75); Tall(X); Tab(95); Weight(X)
    Next X
End Sub

Private Sub cmdSort2_Click() 'this button sorts the info by height
    picResults.Cls
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Tall(Pos + 1) > Tall(Pos) Then
                Temp = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = Temp
                Temp = Jersey(Pos)
                Jersey(Pos) = Jersey(Pos + 1)
                Jersey(Pos + 1) = Temp
                Temp = Experience(Pos)
                Experience(Pos) = Experience(Pos + 1)
                Experience(Pos + 1) = Temp
                Temp = Tall(Pos)
                Tall(Pos) = Tall(Pos + 1)
                Tall(Pos + 1) = Temp
                Temp = Weight(Pos)
                Weight(Pos) = Weight(Pos + 1)
                Weight(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    picResults.Print "Name"; Tab(20); "Height (in Inches)"; Tab(40); "Jersey Number"; Tab(60); "Years of Experience in the NFL"; Tab(95); "Weight (in Pounds)"
    picResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    For X = 1 To 7
        picResults.Print Names(X); Tab(20); Tall(X); Tab(40); Jersey(X); Tab(60); Experience(X); Tab(95); Weight(X)
    Next X
End Sub

Private Sub cmdSort3_Click() 'this button sorts the infor by weight
    picResults.Cls
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If Weight(Pos + 1) > Weight(Pos) Then
                Temp = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = Temp
                Temp = Jersey(Pos)
                Jersey(Pos) = Jersey(Pos + 1)
                Jersey(Pos + 1) = Temp
                Temp = Experience(Pos)
                Experience(Pos) = Experience(Pos + 1)
                Experience(Pos + 1) = Temp
                Temp = Tall(Pos)
                Tall(Pos) = Tall(Pos + 1)
                Tall(Pos + 1) = Temp
                Temp = Weight(Pos)
                Weight(Pos) = Weight(Pos + 1)
                Weight(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    picResults.Print "Name"; Tab(20); "Weight (in Pounds)"; Tab(40); "Jersey Number"; Tab(60); "Years of Experience in the NFL"; Tab(95); "Height (in Inches)"
    picResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    For X = 1 To 7
        picResults.Print Names(X); Tab(20); Weight(X); Tab(40); Jersey(X); Tab(60); Experience(X); Tab(95); Tall(X)
    Next X
End Sub
