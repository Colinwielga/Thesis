VERSION 5.00
Begin VB.Form secondform 
   BackColor       =   &H0080FF80&
   Caption         =   "Form2"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   LinkTopic       =   "Form2"
   ScaleHeight     =   6615
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   3135
      Left            =   2040
      TabIndex        =   8
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox results2 
      BackColor       =   &H00FFC0C0&
      Height          =   3135
      Left            =   2640
      ScaleHeight     =   3075
      ScaleWidth      =   6435
      TabIndex        =   7
      Top             =   3360
      Width           =   6495
   End
   Begin VB.CommandButton gobackbox 
      Caption         =   "Go to First Form"
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Get statistics on Celebrity"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   1815
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H0080FF80&
      Caption         =   "Dr. Phil"
      Height          =   735
      Left            =   6960
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H0080FF80&
      Caption         =   "Princess Di"
      Height          =   735
      Left            =   5520
      MaskColor       =   &H008080FF&
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H0080FF80&
      Caption         =   "Phillis Dillar"
      Height          =   735
      Left            =   3960
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080FF80&
      Caption         =   "Matt Damon"
      Height          =   735
      Left            =   2280
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton option1 
      BackColor       =   &H0080FF80&
      Caption         =   "Garrison Kielor"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Image philimage 
      Height          =   1800
      Left            =   6720
      Picture         =   "Big project2.frx":0000
      Top             =   480
      Width           =   1620
   End
   Begin VB.Image Diimage 
      Height          =   1935
      Left            =   5160
      Picture         =   "Big project2.frx":FE02
      Top             =   360
      Width           =   1440
   End
   Begin VB.Image Dillarimage 
      Height          =   1935
      Left            =   3600
      Picture         =   "Big project2.frx":18F64
      Top             =   360
      Width           =   1440
   End
   Begin VB.Image Mattimage 
      Height          =   1935
      Left            =   2040
      Picture         =   "Big project2.frx":220C6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image keilorimage 
      Height          =   1965
      Left            =   480
      Picture         =   "Big project2.frx":2B42C
      Top             =   360
      Width           =   1500
   End
End
Attribute VB_Name = "secondform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim button As Integer
Dim famousactscores(1 To 11) As Double
Dim famousperson(1 To 11) As String
Dim i As Double




Private Sub Command1_Click()
'read famous people name and ACT score from a file
Open "M:\CS130\Labs\conversionfactors\celbsactscores.txt" For Input As #2
For i = 1 To 11
Input #2, famousperson(i), famousactscores(i)
Next i
' If garrison is picked then information about his life is printed
If button = 1 Then
    results2.Print Tab(1); famousperson(11); Tab(20); "Got a"; famousactscores(11); "on the ACT"
    results2.Print Tab(1); "Occupation:"; Tab(20); "Radio Brodcaster"
    results2.Print Tab(1); "Age:"; Tab(20); "60"
    results2.Print Tab(1); "Weight:"; Tab(20); "183 lbs"
    results2.Print Tab(1); "Favorite Color:"; Tab(20); "Red"
End If
'If matt is picked then information about his life is printed
If button = 2 Then
    results2.Print Tab(1); famousperson(4); Tab(20); "Got a"; famousactscores(4); "on the ACT"
    results2.Print Tab(1); "Occupation:"; Tab(20); "Actor"
    results2.Print Tab(1); "Age:"; Tab(20); "31"
    results2.Print Tab(1); "Weight:"; Tab(20); "169 lbs"
    results2.Print Tab(1); "Favorite Color:"; Tab(20); "Green"
End If
'If phillis is picked then information about her life is printed
If button = 3 Then
    results2.Print Tab(1); famousperson(8); Tab(20); "Got a"; famousactscores(8); "on the ACT"
    results2.Print Tab(1); "Occupation:"; Tab(20); "Comedian"
    results2.Print Tab(1); "Age:"; Tab(20); "Dead"
    results2.Print Tab(1); "Weight:"; Tab(20); "141 lbs"
    results2.Print Tab(1); "Favorite Color:"; Tab(20); "Gold"
End If
'If the princess is picked then information about her life is printed
If button = 4 Then
    results2.Print Tab(1); famousperson(10); Tab(20); "Got a"; famousactscores(10); "on the ACT"
    results2.Print Tab(1); "Occupation:"; Tab(20); "Princess"
    results2.Print Tab(1); "Age:"; Tab(20); "Dead"
    results2.Print Tab(1); "Weight:"; Tab(20); "132 lbs"
    results2.Print Tab(1); "Favorite Color:"; Tab(20); "Red"
End If
'If phil is picked then information about his life is printed
If button = 5 Then
    results2.Print Tab(1); famousperson(7); Tab(20); "Got a"; famousactscores(7); "on the ACT"
    results2.Print Tab(1); "Occupation:"; Tab(20); "Dr. and TV show host"
    results2.Print Tab(1); "Age:"; Tab(20); "54"
    results2.Print Tab(1); "Weight:"; Tab(20); "2078 lbs"
    results2.Print Tab(1); "Favorite Color:"; Tab(20); "Teal"
End If
Close #2
End Sub

Private Sub Command2_Click()
results2.Cls
End Sub

Private Sub gobackbox_Click()
FirstForm.Show
secondform.Hide
End Sub

Private Sub option1_Click()
button = 1
End Sub

Private Sub Option2_Click()
button = 2
End Sub

Private Sub Option3_Click()
button = 3
End Sub

Private Sub Option4_Click()
button = 4
End Sub

Private Sub Option5_Click()
button = 5
End Sub
