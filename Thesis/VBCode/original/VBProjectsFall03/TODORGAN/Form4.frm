VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FF0000&
   Caption         =   "Form4"
   ClientHeight    =   5925
   ClientLeft      =   2790
   ClientTop       =   3210
   ClientWidth     =   9990
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form4"
   ScaleHeight     =   5925
   ScaleWidth      =   9990
   Begin VB.CommandButton Command2 
      Caption         =   "Return to main screen"
      Height          =   735
      Left            =   240
      TabIndex        =   13
      Top             =   5040
      Width           =   1335
   End
   Begin VB.PictureBox Results 
      BackColor       =   &H000080FF&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   7395
      TabIndex        =   12
      Top             =   3120
      Width           =   7455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find Total"
      Height          =   1215
      Left            =   5880
      TabIndex        =   11
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Quantity 
      Height          =   615
      Left            =   5040
      TabIndex        =   10
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox Auto 
      Height          =   615
      Left            =   5040
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Please enter how many of these vehicles you would like"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   4695
   End
   Begin VB.Label Label8 
      Caption         =   "Please enter the number of the vehicle you would like to purchase"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Label Label7 
      Caption         =   "6. Mercedes Benz AMG G55"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "4. Hummer H2"
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "3. Cadillac Escalade"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "7. Land Rover Range Rover"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "5. Lincoln Navigator"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "2. BMW X5 4.6is"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "1. Acura MDX"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Buying an SUV
'Form Name : Form4.frm
'Author: Tom Dorgan
'Date Written: October 28, 2003
'Purpose of Form: This form helps the user price out an SUV, it is useful for the user
                    ' who is looking for more than one of the same vehicle.




Dim A As Single
Dim Q As Single
Dim T As Single
Dim V(1 To 7), D(1 To 7) As String
Dim P(1 To 7), L(1 To 7), H(1 To 7), C(1 To 7), F(1 To 7), S(1 To 7) As Single
Dim i As Integer
Public Path As String



Private Sub Command1_Click()
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1

A = Auto.Text
Q = Quantity.Text
T = P(A) * Q
Results.Print Q; V(A); " is "; FormatCurrency(T, 0)
Results.Print "Now visit our store and drive home in a new luxury SUV!"


End Sub

Private Sub Command2_Click()
Form3.Hide
Form1.Show
Form2.Hide
Form4.Hide

End Sub

Private Sub Form_Load()
Path = "N:\cs130\handin\TODORGAN\"

End Sub
