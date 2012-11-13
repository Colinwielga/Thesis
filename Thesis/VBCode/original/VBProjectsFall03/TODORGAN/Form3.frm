VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form3"
   ClientHeight    =   8760
   ClientLeft      =   2640
   ClientTop       =   1680
   ClientWidth     =   10485
   FillColor       =   &H0000FFFF&
   LinkTopic       =   "Form3"
   ScaleHeight     =   8760
   ScaleWidth      =   10485
   Begin VB.CommandButton Price 
      Caption         =   "Price"
      Height          =   495
      Left            =   480
      TabIndex        =   24
      Top             =   3360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Engine 
      Caption         =   "Engine Size"
      Height          =   495
      Left            =   480
      TabIndex        =   23
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Power 
      Caption         =   "Horse Power"
      Height          =   495
      Left            =   480
      TabIndex        =   22
      Top             =   4560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton City 
      Caption         =   "Fuel Economy:City"
      Height          =   495
      Left            =   480
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Highway 
      Caption         =   "Fuel Economy:Highway"
      Height          =   495
      Left            =   480
      TabIndex        =   20
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Seating 
      Caption         =   "Seating Capacity"
      Height          =   495
      Left            =   480
      TabIndex        =   19
      Top             =   6360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Return to main screen"
      Height          =   855
      Left            =   360
      TabIndex        =   18
      Top             =   7800
      Width           =   2775
   End
   Begin VB.PictureBox Picture7 
      Height          =   2295
      Left            =   3840
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   3315
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Picture6 
      Height          =   2295
      Left            =   3840
      Picture         =   "Form3.frx":27DF
      ScaleHeight     =   2235
      ScaleWidth      =   3315
      TabIndex        =   16
      Top             =   1080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Picture5 
      Height          =   2295
      Left            =   3840
      Picture         =   "Form3.frx":5E6B
      ScaleHeight     =   2235
      ScaleWidth      =   3315
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Picture4 
      Height          =   2295
      Left            =   3840
      Picture         =   "Form3.frx":9A1A
      ScaleHeight     =   2235
      ScaleWidth      =   3315
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Picture3 
      Height          =   2295
      Left            =   3840
      Picture         =   "Form3.frx":DA93
      ScaleHeight     =   2235
      ScaleWidth      =   3315
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   3840
      Picture         =   "Form3.frx":10033
      ScaleHeight     =   2235
      ScaleWidth      =   3315
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Specs 
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   3840
      ScaleHeight     =   3915
      ScaleWidth      =   6315
      TabIndex        =   11
      Top             =   3720
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   3840
      Picture         =   "Form3.frx":134DD
      ScaleHeight     =   2235
      ScaleWidth      =   3315
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Rover 
      Caption         =   "Land Rover Range Rover"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   6960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton X5 
      Caption         =   "BMW X5 4.6is"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Escalade 
      Caption         =   "Cadillac Escalade"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton H2 
      Caption         =   "Hummer H2"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Navigator 
      Caption         =   "Lincoln Navigator"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton AMG 
      Caption         =   "Mercedes Benz AMG G55"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton MDX 
      Caption         =   "Acura MDX"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Compare an aspect of a vehicle"
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View a vehicle and its specifications"
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "You have chosen Luxury SUVs.  Please select the load button and then your option below."
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Buying an SUV
'Form Name : Form3.frm
'Author: Tom Dorgan
'Date Written: October 28, 2003
'Purpose of Form: To rank different convenience and comfort aspects of an SUV.
                  'Also to view an SUV and different specifications about it.









Dim V(1 To 7), D(1 To 7) As String
Dim P(1 To 7), L(1 To 7), H(1 To 7), C(1 To 7), F(1 To 7), S(1 To 7) As Single
Dim i As Integer
Public Path As String


Private Sub AMG_Click()
Specs.Cls
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1

Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = True
Picture7.Visible = False
Specs.Print V(6)
Specs.Print "Price"; Tab(30); FormatCurrency(P(6), 0)
Specs.Print "Engine Size"; Tab(30); L(6); " liter"
Specs.Print "Horsepower"; Tab(30); H(6)
Specs.Print "City Fuel Economy"; Tab(30); C(6)
Specs.Print "Highway Fuel Economy"; Tab(30); F(6)
Specs.Print "Type of Vehicle"; Tab(30); D(6)
Specs.Print "Passenger Capacity"; Tab(30); S(6)
End Sub

Private Sub City_Click()
Specs.Cls
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1

Dim pass As Integer, comp As Integer, z As Integer
Dim temp As Single
Dim N As Integer
N = 7
Dim tempV As String
For pass = 1 To N - 1
    For z = 1 To N - pass
        If C(z) > C(z + 1) Then
            temp = C(z + 1)
            C(z + 1) = C(z)
            C(z) = temp
            tempV = V(z + 1)
            V(z + 1) = V(z)
            V(z) = tempV
          End If
        Next z
Next pass
    For i = 1 To 7
    Specs.Print V(i); Tab(30); C(i); " mpg in the city"
Next i
End Sub

Private Sub Command1_Click()
MDX.Visible = True
X5.Visible = True
Escalade.Visible = True
H2.Visible = True
Navigator.Visible = True
AMG.Visible = True
Rover.Visible = True
Price.Visible = False
Engine.Visible = False
Power.Visible = False
City.Visible = False
Highway.Visible = False
Seating.Visible = False

End Sub

Private Sub Command2_Click()
Specs.Cls
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Price.Visible = True
Engine.Visible = True
Power.Visible = True
City.Visible = True
Highway.Visible = True
Seating.Visible = True
MDX.Visible = False
X5.Visible = False
Escalade.Visible = False
H2.Visible = False
Navigator.Visible = False
AMG.Visible = False
Rover.Visible = False

End Sub

Private Sub Command3_Click()
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1
End Sub

Private Sub Command4_Click()
Form3.Hide
Form1.Show
Form2.Hide
Form4.Hide


End Sub

Private Sub Engine_Click()
Specs.Cls
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1

Dim pass As Integer, comp As Integer, z As Integer
Dim temp As Single
Dim N As Integer
N = 7
Dim tempV As String
For pass = 1 To N - 1
    For z = 1 To N - pass
        If L(z) > L(z + 1) Then
            temp = L(z + 1)
            L(z + 1) = L(z)
            L(z) = temp
            tempV = V(z + 1)
            V(z + 1) = V(z)
            V(z) = tempV
          End If
        Next z
Next pass
For i = 1 To 7
    Specs.Print V(i); Tab(30); L(i); " liters"
Next i
End Sub

Private Sub Escalade_Click()
Specs.Cls
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1

Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = True
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Specs.Print V(3)
Specs.Print "Price"; Tab(30); FormatCurrency(P(3), 0)
Specs.Print "Engine Size"; Tab(30); L(3); " liter"
Specs.Print "Horsepower"; Tab(30); H(3)
Specs.Print "City Fuel Economy"; Tab(30); C(3)
Specs.Print "Highway Fuel Economy"; Tab(30); F(3)
Specs.Print "Type of Vehicle"; Tab(30); D(3)
Specs.Print "Passenger Capacity"; Tab(30); S(3)
End Sub

Private Sub Form_Load()
Path = "N:\cs130\handin\TODORGAN\"
End Sub

Private Sub H2_Click()
Specs.Cls
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1

Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = True
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Specs.Print V(4)
Specs.Print "Price"; Tab(30); FormatCurrency(P(4), 0)
Specs.Print "Engine Size"; Tab(30); L(4); " liter"
Specs.Print "Horsepower"; Tab(30); H(4)
Specs.Print "City Fuel Economy"; Tab(30); C(4)
Specs.Print "Highway Fuel Economy"; Tab(30); F(4)
Specs.Print "Type of Vehicle"; Tab(30); D(4)
Specs.Print "Passenger Capacity"; Tab(30); S(4)
End Sub

Private Sub Highway_Click()
Specs.Cls
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1

Dim pass As Integer, comp As Integer, z As Integer
Dim temp As Single
Dim N As Integer
N = 7
Dim tempV As String
For pass = 1 To N - 1
    For z = 1 To N - pass
        If F(z) > F(z + 1) Then
            temp = F(z + 1)
            F(z + 1) = F(z)
            F(z) = temp
            tempV = V(z + 1)
            V(z + 1) = V(z)
            V(z) = tempV
          End If
        Next z
Next pass
For i = 1 To 7
    Specs.Print V(i); Tab(30); F(i); " mpg on the highway"
Next i
End Sub

Private Sub MDX_Click()
Specs.Cls
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1

Picture1.Visible = True
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Specs.Print V(1)
Specs.Print "Price"; Tab(30); FormatCurrency(P(1), 0)
Specs.Print "Engine Size"; Tab(30); L(1); " liter"
Specs.Print "Horsepower"; Tab(30); H(1)
Specs.Print "City Fuel Economy"; Tab(30); C(1)
Specs.Print "Highway Fuel Economy"; Tab(30); F(1)
Specs.Print "Type of Vehicle"; Tab(30); D(1)
Specs.Print "Passenger Capacity"; Tab(30); S(1)
End Sub

Private Sub Navigator_Click()
Specs.Cls
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1

Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = True
Picture6.Visible = False
Picture7.Visible = False
Specs.Print V(5)
Specs.Print "Price"; Tab(30); FormatCurrency(P(5), 0)
Specs.Print "Engine Size"; Tab(30); L(5); " liter"
Specs.Print "Horsepower"; Tab(30); H(5)
Specs.Print "City Fuel Economy"; Tab(30); C(5)
Specs.Print "Highway Fuel Economy"; Tab(30); F(5)
Specs.Print "Type of Vehicle"; Tab(30); D(5)
Specs.Print "Passenger Capacity"; Tab(30); S(5)
End Sub

Private Sub Power_Click()
Specs.Cls
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1

Dim pass As Integer, comp As Integer, z As Integer
Dim temp As Single
Dim N As Integer
N = 7
Dim tempV As String
For pass = 1 To N - 1
    For z = 1 To N - pass
        If H(z) > H(z + 1) Then
            temp = H(z + 1)
            H(z + 1) = H(z)
            H(z) = temp
            tempV = V(z + 1)
            V(z + 1) = V(z)
            V(z) = tempV
          End If
        Next z

Next pass
For i = 1 To 7
    Specs.Print V(i); Tab(30); H(i); " horsepower"
Next i

End Sub

Private Sub Price_Click()
Specs.Cls
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1

Dim pass As Integer, comp As Integer, iz As Integer
Dim temp As Single
Dim N As Integer
N = 7
Dim tempV As String
For pass = 1 To N - 1
    For z = 1 To N - pass
        If P(z) < P(z + 1) Then
            temp = P(z + 1)
            P(z + 1) = P(z)
            P(z) = temp
            tempV = V(z + 1)
            V(z + 1) = V(z)
            V(z) = tempV
          End If
        Next z

Next pass
For i = 1 To 7
    Specs.Print V(i); Tab(30); FormatCurrency(P(i), 0)
Next i
End Sub

Private Sub Rover_Click()
Specs.Cls
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1

Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = True
Specs.Print V(7)
Specs.Print "Price"; Tab(30); FormatCurrency(P(7), 0)
Specs.Print "Engine Size"; Tab(30); L(7); " liter"
Specs.Print "Horsepower"; Tab(30); H(7)
Specs.Print "City Fuel Economy"; Tab(30); C(7)
Specs.Print "Highway Fuel Economy"; Tab(30); F(7)
Specs.Print "Type of Vehicle"; Tab(30); D(7)
Specs.Print "Passenger Capacity"; Tab(30); S(7)

End Sub

Private Sub Seating_Click()
Specs.Cls
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1

Dim pass As Integer, comp As Integer, z As Integer
Dim temp As Single
Dim N As Integer
N = 7
Dim tempV As String
Specs.Print "Maximum Seating Capacity:"
For pass = 1 To N - 1
    For z = 1 To N - pass
        If S(z) > S(z + 1) Then
            temp = S(z + 1)
            S(z + 1) = S(z)
            S(z) = temp
            tempV = V(z + 1)
            V(z + 1) = V(z)
            V(z) = tempV
          End If
        Next z

Next pass
For i = 1 To 7
    Specs.Print V(i); Tab(30); S(i); " passengers"
Next i
End Sub

Private Sub X5_Click()
Specs.Cls
Open Path & "SUVs.txt" For Input As #1
For i = 1 To 7
    Input #1, V(i), P(i), L(i), H(i), C(i), F(i), D(i), S(i)
Next i
Close #1

Picture1.Visible = False
Picture2.Visible = True
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Specs.Print V(2)
Specs.Print "Price"; Tab(30); FormatCurrency(P(2), 0)
Specs.Print "Engine Size"; Tab(30); L(2); " liter"
Specs.Print "Horsepower"; Tab(30); H(2)
Specs.Print "City Fuel Economy"; Tab(30); C(2)
Specs.Print "Highway Fuel Economy"; Tab(30); F(2)
Specs.Print "Type of Vehicle"; Tab(30); D(2)
Specs.Print "Passenger Capacity"; Tab(30); S(2)

End Sub
