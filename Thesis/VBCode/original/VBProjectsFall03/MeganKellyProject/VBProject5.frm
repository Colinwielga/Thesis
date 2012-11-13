VERSION 5.00
Begin VB.Form vitals5 
   BackColor       =   &H00C0C000&
   Caption         =   "I know, I know..."
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Programmed by Megan Kelly"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   2040
      Width           =   2655
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
Dim BMI As Single

Private Sub Continue5_Click()
' Which not so nice person would you like to beat up today? "Megan'sVBProject.vbp"
'                       Intro1 (VBProject4.frm)
'                       Megan Kelly 11/03/03
' Purpose:  The purpose of this form is to collect information about the person and calculate it into a running sum to help determine the completely unscientific outcome of this exercise.

vitals5.Visible = False
'Input an individual's weight

If weight > 0 Then
weight = Int(InputBox("Input Weight (in pounds)", "Weight"))
Else
    MsgBox ("Please enter your weight in pounds.")
    weight = Int(InputBox("Input Weight (in pounds)", "Weight"))
If weight > 100 Then
    If weight >= 100 Then sum = sum
    If weight >= 120 Then sum = sum + 1
    If weight >= 140 Then sum = sum + 1
    If weight >= 160 Then sum = sum + 1
    If weight >= 180 Then sum = sum + 1
    If weight >= 200 Then sum = sum + 1
    If weight >= 250 Then sum = sum + 2
    If weight <= 100 Then sum = sum - 3
End If
'Input an individual's age
If age > 0 Then
    age = InputBox("Input Age", "Age")
ElseIf age <= 0 Then
    MsgBox ("Please enter your real age.")
    age = InputBox("Input Age", "Age")
End If

If age < 15 Then sum = sum - 3
If age >= 15 Then sum = sum + 1
If age >= 17 Then sum = sum + 1
If age >= 19 Then sum = sum + 1
If age >= 21 Then sum = sum + 1
If age >= 45 Then sum = sum - 1
If age >= 50 Then sum = sum - 1
If age >= 60 Then sum = sum - 1
If age >= 70 Then sum = sum - 3
End If

'Input an individual's height

If height < 0 Then
    height = Int(InputBox("Input Height (in inches)", "Height"))
Else
    MsgBox ("Please enter your correct height in inches.")
    height = Int(InputBox("Input Height (in inches)", "Height"))
End If
If height < 48 Then sum = sum - 3
If height >= 48 Then
    If height >= 48 Then sum = sum
    If height >= 60 Then sum = sum + 1
    If height >= 72 Then sum = sum + 2
    If height >= 84 Then sum = sum + 4
Else: If height <= 24 Then MsgBox ("Please enter your correct height in inches.")
End If

'Calculate the user's Body Mass Index
BMI = (weight / 2.2) / (height / 39.4) ^ 2
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
