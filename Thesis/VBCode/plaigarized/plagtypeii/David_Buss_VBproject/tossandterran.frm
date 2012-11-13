VERSION 5.00
Begin VB.Form one
   Caption         =   "one"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "one"
   Picture         =   "toss and terran.frx":0000
   ScaleHeight     =   5565
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.PictureBox PicName
      Height          =   495
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command4
      Caption         =   "Enter"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command3
      Caption         =   "For Protoss"
      Height          =   975
      Left            =   2760
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command2
      Caption         =   "For Terrans"
      Height          =   975
      Left            =   2640
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command1
      Caption         =   "What Race Am I?"
      Height          =   975
      Left            =   2640
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2
      Caption         =   "<-------------------------------- Enter Your Name Here----------------------------"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Label Label1
      Caption         =   "Start with a quick quiz!"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "one"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub stuffone_Click()


Dim aaa As Integer, bbb As String, ccc As String
Dim ddd As Integer, eee As String
Dim fff As Boolean

fff = False
Do Until fff = True

    bbb = LCase(InputBox("What color is more awesome; Blue or Purple?", "Blue Or Purple"))
    ccc = LCase(InputBox("What weapon is more awesome; Guns or Laser Beams?", "Guns or Lasers"))
    eee = LCase(InputBox("What is just more awesome; Giant freaking robot or Psionic Storms ?", "Robot or Storms"))

    If bbb = "blue" Then
        ddd = ddd + 1
    End If

    If bbb = "purple" Then
        aaa = aaa + 1
    End If

    If ccc = "guns" Then
        ddd = ddd + 1
    End If

    If ccc = "lasers" Then
        aaa = aaa + 1
    End If

    If eee = "robot" Then
        ddd = ddd + 1
    End If

    If bbb = "storms" Then
        aaa = aaa + 1
    End If

    If aaa > ddd Then
        fff = True
        MsgBox ("You're totally a 'Brotoss' Protoss player!")
        Command3.Visible = True
        Command2.Visible = False

    End If

    If ddd > aaa Then
        fff = True
        MsgBox ("You're totally the best race, Terran!")
        Command3.Visible = False
        Command2.Visible = True
    End If

    If fff = False Then
        MsgBox ("Try again,please!")

    End If
Loop
Label1.Visible = False
End Sub

Private Sub stufftwo_Click()
one.Hide
two.Show
three.Hide

End Sub

Private Sub stuffthree_Click()
one.Hide
two.Hide
three.Show

End Sub

Private Sub stufffour_Click()

If Text1 = "" Then
        MsgBox ("Please enter your name")
Else
    myname = Text1
    PicName.Print myname
    Command1.Visible = True
    Text1.Visible = False
    Command4.Visible = False
    Label1.Visible = True
    Label2.Visible = False
End If
End Sub

Private Sub Form_Load()
one.Show
two.Hide
three.Hide

End Sub

