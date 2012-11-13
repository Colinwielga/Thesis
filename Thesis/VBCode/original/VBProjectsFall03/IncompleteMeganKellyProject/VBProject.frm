VERSION 5.00
Begin VB.Form vitals5 
   BackColor       =   &H00C0C000&
   Caption         =   "Form5"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4890
   LinkTopic       =   "Form5"
   ScaleHeight     =   2400
   ScaleWidth      =   4890
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Quit5 
      BackColor       =   &H00C0C000&
      Caption         =   "Quit"
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Continue5 
      BackColor       =   &H00C0C000&
      Caption         =   "Continue"
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Just a few more questions....  would you like to keep going?"
      Height          =   1335
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "vitals5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Continue5_Click()
vitals5.Visible = False
Dim BMI As Single
Dim height As Single
Dim weight As Single
height = Int(InputBox("Input Height (in inches)", "Height"))
weight = Int(InputBox("Input Weight (in pounds)", "Weight"))
BMI = (weight / 2.2) / (height / 39.4) ^ 2
age = InputBox("Input Age", "Age")

If weight > 100 Then
    If weight >= 100 Then sum = sum
    If weight >= 120 Then sum = sum + 1
    If weight >= 140 Then sum = sum + 1
    If weight >= 160 Then sum = sum + 1
    If weight >= 180 Then sum = sum + 1
    If weight >= 200 Then sum = sum + 1
    If weight >= 250 Then sum = sum + 2
    If weight <= 100 Then sum = sum - 3
    ElseIf weight <= 50 Then MsgBox ("Sorry, but you need to enter a weight greater than 50 lbs.")
End If

If age <= 0 Then MsgBox ("Please enter your real age.")
If age < 15 Then sum = sum - 3
If age >= 15 Then sum = sum + 1
If age >= 17 Then sum = sum + 1
If age >= 19 Then sum = sum + 1
If age >= 21 Then sum = sum + 1
If age >= 45 Then sum = sum - 1
If age >= 50 Then sum = sum - 1
If age >= 60 Then sum = sum - 1
If age >= 70 Then sum = sum - 3

If height < 48 Then sum = sum - 3
If height >= 48 Then
    If height >= 48 Then sum = sum
    If height >= 60 Then sum = sum + 1
    If height >= 72 Then sum = sum + 2
    If height >= 84 Then sum = sum + 4
Else: If height <= 24 Then MsgBox ("Please enter your correct height in inches.")
End If


Select Case BMI
    Case Is < 19
        sum = sum - 7
    Case 19 To 25
        sum = sum + 5
    Case 25.01 To 30
        sum = sum - 1
    Case 30.01 To 40
        sum = sum - 5
    Case Is > 40
        sum = sum - 7
End Select

thirddegree6.Visible = True

End Sub


Private Sub Quit5_Click()
End
End Sub
