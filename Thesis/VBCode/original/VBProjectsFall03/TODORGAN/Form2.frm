VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H000000FF&
   Caption         =   "Form2"
   ClientHeight    =   5220
   ClientLeft      =   3255
   ClientTop       =   3210
   ClientWidth     =   9330
   FillColor       =   &H00C00000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   5220
   ScaleWidth      =   9330
   Begin VB.CommandButton Command4 
      Caption         =   "Return to main screen"
      Height          =   975
      Left            =   3360
      TabIndex        =   4
      Top             =   3720
      Width           =   3735
   End
   Begin VB.PictureBox Results 
      BackColor       =   &H00FF0000&
      Height          =   2655
      Left            =   3120
      ScaleHeight     =   2595
      ScaleWidth      =   5355
      TabIndex        =   3
      Top             =   480
      Width           =   5415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "   Need an SUV that will            handle your oversized                       family?                   Click here"
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Looking for an SUV that has enough horsepower to handle your lead foot?    Click here"
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Are you someone who is looking for an SUV with certan fuel economy?     Click here"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Buying an SUV
'Form Name : Form2.frm
'Author: Tom Dorgan
'Date Written: October 28, 2003
'Purpose of Form: To help the user decide what SUV is right for him, this site personalizes
                    'a car according to the user's recommendation in fuel economy,
                    'horsepower, and seating capacity.



Dim V(1 To 7), D(1 To 7) As String
Dim P(1 To 7), L(1 To 7), H(1 To 7), C(1 To 7), F(1 To 7), S(1 To 7) As Single
Dim i As Integer
Dim fuel As Single
Dim HP As Single
Dim seat As Single

Public Path As String



Private Sub Command1_Click()
Results.Cls

Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1
fuel = InputBox("Please enter your minimum fuel economy on the highway")
For i = 1 To 7
    If fuel <= F(i) Then
        Results.Print V(i); " has "; C(i); " mpg in the city and "; F(i); " mpg on the highway."
    End If
Next i
End Sub

Private Sub Command2_Click()
Results.Cls

Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1
HP = InputBox("Please enter your minimum desired horsepower")
For i = 1 To 7
    If HP <= H(i) Then
        Results.Print V(i); " has "; H(i); " horsepower."
    End If
Next i

End Sub

Private Sub Command3_Click()
Results.Cls
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1
seat = InputBox("Please enter your minimum desired seating capacity")
For i = 1 To 7
    If seat <= S(i) Then
        Results.Print V(i); " has room for "; S(i); " passengers."
    End If
Next i

End Sub

Private Sub Command4_Click()
Form3.Hide
Form1.Show
Form2.Hide
Form4.Hide

End Sub

Private Sub Form_Load()
Path = "N:\cs130\handin\TODORGAN\"

End Sub
