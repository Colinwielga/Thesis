VERSION 5.00
Begin VB.Form frmCal 
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSpped 
      Height          =   855
      Left            =   7080
      TabIndex        =   10
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdManual 
      Caption         =   "Manual Calculate"
      Height          =   1095
      Left            =   2640
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtTime 
      Height          =   735
      Left            =   7080
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear11 
      Caption         =   "Clear"
      Height          =   855
      Left            =   4680
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdBackc 
      Caption         =   "Back"
      Height          =   1095
      Left            =   7200
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picResults4 
      Height          =   2415
      Left            =   600
      ScaleHeight     =   2355
      ScaleWidth      =   3675
      TabIndex        =   2
      Top             =   3360
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Weapon Calculator"
      Height          =   1095
      Left            =   600
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display the power of weapon"
      Height          =   1095
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblspped 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   6240
      TabIndex        =   9
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lbltime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   6240
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "Distance = speed x time"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1095
      Left            =   3480
      TabIndex        =   5
      Top             =   720
      Width           =   3495
   End
   Begin VB.Image imgWeapon 
      Height          =   5985
      Left            =   0
      Picture         =   "frmCal.frx":0000
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBackc_Click()
frmCal.Hide
frmMain.Show
End Sub

Private Sub cmdClear11_Click()
picResults4.Cls
End Sub

Private Sub cmdDisplay_Click()
Dim sname(1 To 100) As String
Dim sspeed(1 To 100) As Single
Dim ctr As Integer
Dim pos As Integer

    Open App.Path & "\weaponpower.txt" For Input As #1
    ctr = 0
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, sname(ctr), sspeed(ctr)
    Loop
    Close #1
        picResults4.Print "Name", "speed of bullet or weapon"
        picResults4.Print "************************************"
     For pos = 1 To ctr
        picResults4.Print sname(pos), sspeed(pos)
    Next pos
End Sub

Private Sub cmdManual_Click()
Dim time As Single
Dim speed As Single
Dim distance As Single
time = txtTime.Text
speed = txtSpped.Text
distance = time * speed
picResults4.Print distance
End Sub

Private Sub Command2_Click()
Shell "calc.exe"
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtTime_Change()

End Sub
