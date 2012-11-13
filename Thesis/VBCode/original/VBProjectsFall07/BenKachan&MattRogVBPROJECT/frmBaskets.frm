VERSION 5.00
Begin VB.Form frmBaskets 
   BackColor       =   &H80000014&
   Caption         =   "Basketball Shoot Around"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmBaskets.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate your shooting percentage and total score!"
      Height          =   1455
      Left            =   8640
      TabIndex        =   7
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Menu"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton cmdHalf 
      Caption         =   "Half Court Shot!"
      Height          =   615
      Left            =   3720
      TabIndex        =   5
      Top             =   7680
      Width           =   2775
   End
   Begin VB.CommandButton cmdThree 
      Caption         =   "The Three Ball!"
      Height          =   1095
      Left            =   1200
      TabIndex        =   4
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdMid 
      Caption         =   "Mid Range Jumper"
      Height          =   855
      Left            =   7440
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCorner 
      Caption         =   "Corner Three!"
      Height          =   1575
      Left            =   9960
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton cmdFree 
      Caption         =   "Free Throw Attempt"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdLayup 
      Caption         =   "Easy Lay-up"
      Height          =   735
      Left            =   5880
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "frmBaskets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sum As Integer, CTR As Integer, tota As Integer

Private Sub cmdCalc_Click()
    'Here we display the total number of points and calculate and display average
    Dim avg As Single
    avg = tota / CTR
    MsgBox "You managed to score " & sum & " points while shooting " & FormatPercent(avg)
End Sub

Private Sub cmdCorner_Click()
    'We simply ask the user to input power and height of shot and then the numbers
    'are put into our equation and it is determined whether or not the player makes the shot
    Dim power As Integer, height As Integer, result As Single
    power = InputBox("How HARD would you like to shoot 1-10? 10 being the hardest")
    height = InputBox("How HIGH would you like to shoot 1-10? 10 being the highest")
    result = (power * height) / 10
    Select Case result
        Case 6 To 7.5
        MsgBox "Corner Three! He's On Fire!"
        sum = sum + 3
        CTR = CTR + 1
        tota = tota + 1
        Case Else
        MsgBox "Is that hoop regulation size or what? you missed"
        CTR = CTR + 1
    End Select
End Sub

Private Sub cmdFree_Click()
    'We simply ask the user to input power and height of shot and then the numbers
    'are put into our equation and it is determined whether or not the player makes the shot
    Dim power As Integer, height As Integer, result As Single
    power = InputBox("How HARD would you like to shoot 1-10? 10 being the hardest")
    height = InputBox("How HIGH would you like to shoot 1-10? 10 being the highest")
    result = (power * height) / 10
    Select Case result
        Case 2 To 5
        MsgBox "You made the free throw congratulations"
        sum = sum + 1
        CTR = CTR + 1
        tota = tota + 1
        Case Else
        MsgBox "Is that hoop regulation size or what? you missed"
        CTR = CTR + 1
    End Select
End Sub

Private Sub cmdHalf_Click()
    Dim power As Integer, height As Integer, result As Single
    power = InputBox("How HARD would you like to shoot 1-10? 10 being the hardest")
    height = InputBox("How HIGH would you like to shoot 1-10? 10 being the highest")
    result = (power * height) / 10
    If (result = 8) Then
        MsgBox "Do you believe in miracles? YES!!"
        sum = sum + 3
        CTR = CTR + 1
        tota = tota + 1
    Else
        MsgBox "Is that hoop regulation size or what? you missed"
        CTR = CTR + 1
    End If
    
End Sub

Private Sub cmdLayup_Click()
    'We simply ask the user to input power and height of shot and then the numbers
    'are put into our equation and it is determined whether or not the player makes the shot
    'also the total number of points and shooting percentage are calculated upon user request
    Dim power As Integer, height As Integer, result As Single
    power = InputBox("How HARD would you like to shoot 1-10? 10 being the hardest")
    height = InputBox("How HIGH would you like to shoot 1-10? 10 being the highest")
    result = (power * height) / 10
    Select Case result
        Case 1 To 3.5
        MsgBox "Easy Lay-up. why don't you challenge yourself?"
        sum = sum + 2
        CTR = CTR + 1
        tota = tota + 1
        Case Else
        MsgBox "Is that hoop regulation size or what? you missed"
        CTR = CTR + 1
    End Select
End Sub

Private Sub cmdMid_Click()
    'We simply ask the user to input power and height of shot and then the numbers
    'are put into our equation and it is determined whether or not the player makes the shot
    Dim power As Integer, height As Integer, result As Single
    power = InputBox("How HARD would you like to shoot 1-10? 10 being the hardest")
    height = InputBox("How HIGH would you like to shoot 1-10? 10 being the highest")
    result = (power * height) / 10
    Select Case result
        Case 3.5 To 5.25
        MsgBox "You made the mid ranger. BALLIN!"
        sum = sum + 2
        CTR = CTR + 1
        tota = tota + 1
        Case Else
        MsgBox "Is that hoop regulation size or what? you missed"
        CTR = CTR + 1
    End Select
End Sub

Private Sub cmdReturn_Click()
    frmBaskets.Hide
    frmHome.Show
End Sub

Private Sub cmdThree_Click()
    'We simply ask the user to input power and height of shot and then the numbers
    'are put into our equation and it is determined whether or not the player makes the shot
    Dim power As Integer, height As Integer, result As Single
    power = InputBox("How HARD would you like to shoot 1-10? 10 being the hardest")
    height = InputBox("How HIGH would you like to shoot 1-10? 10 being the highest")
    result = (power * height) / 10
    Select Case result
        Case 5.5 To 7.25
        MsgBox "Nothing but net! Put it on the board... YES!"
        sum = sum + 3
        CTR = CTR + 1
        tota = tota + 1
        Case Else
        MsgBox "Is that hoop regulation size or what? you missed"
        CTR = CTR + 1
    End Select
End Sub
