VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00400000&
   Caption         =   "Weights"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form8"
   ScaleHeight     =   8865
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfindweight 
      Caption         =   "Click to find the weights of the 3 models closest to your weight preference"
      Enabled         =   0   'False
      Height          =   1575
      Left            =   600
      TabIndex        =   4
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CommandButton cmdsortweight 
      Caption         =   "Click to list the weights from greatest to smallest"
      Enabled         =   0   'False
      Height          =   1815
      Left            =   600
      TabIndex        =   3
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click to return to the main menu"
      Height          =   1815
      Left            =   720
      TabIndex        =   2
      Top             =   6840
      Width           =   2775
   End
   Begin VB.CommandButton cmdweight 
      Caption         =   "Click to see the weight for all of the models"
      Height          =   2055
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H8000000E&
      Height          =   6255
      Left            =   4680
      ScaleHeight     =   6195
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Names(1 To 23) As String
Dim Prices(1 To 23) As Integer
Dim Weight(1 To 23) As Single
Dim Milespergallon(1 To 23) As Integer
Dim Cubicinch(1 To 23) As Single
Dim I As Integer

Private Sub cmdfindweight_Click()
picresults.Cls
picresults.Print "Names of the 3 models"; Tab(40); "Weights(Lbs.) of the 3 bikes"
picresults.Print "********************************************************************************************************************************************************"
usersinput = InputBox("enter a weight you prefer")
For Pass = 1 To 22
    For I = 1 To 23 - Pass
        If Abs(Weight(I) - usersinput) >= Abs(Weight(I + 1) - usersinput) Then
        Temp = Names(I)
        Names(I) = Names(I + 1)
        Names(I + 1) = Temp
        Temp = Prices(I)
        Prices(I) = Prices(I + 1)
        Prices(I + 1) = Temp
        Temp = Weight(I)
        Weight(I) = Weight(I + 1)
        Weight(I + 1) = Temp
        Temp = Milespergallon(I)
        Milespergallon(I) = Milespergallon(I + 1)
        Milespergallon(I + 1) = Temp
        Temp = Cubicinch(I)
        Cubicinch(I) = Cubicinch(I + 1)
        Cubicinch(I + 1) = Temp
        End If
    Next I
Next Pass
For I = 1 To 3
    picresults.Print Names(I); Tab(50); Weight(I); "Lbs."
Next I
I = 1
MsgBox ("The model with the weight closest to your preference is " & Names(I))
End Sub

Private Sub cmdsortweight_Click()
picresults.Cls
picresults.Print "Names"; Tab(40); "Weights(Lbs.) from greatest to smallest"
picresults.Print "****************************************************************************************************************************************************************"
For Pass = 1 To 22
    For I = 1 To 23 - Pass
        If Weight(I) <= Weight(I + 1) Then
        Temp = Names(I)
        Names(I) = Names(I + 1)
        Names(I + 1) = Temp
        Temp = Prices(I)
        Prices(I) = Prices(I + 1)
        Prices(I + 1) = Temp
        Temp = Weight(I)
        Weight(I) = Weight(I + 1)
        Weight(I + 1) = Temp
        Temp = Milespergallon(I)
        Milespergallon(I) = Milespergallon(I + 1)
        Milespergallon(I + 1) = Temp
        Temp = Cubicinch(I)
        Cubicinch(I) = Cubicinch(I + 1)
        Cubicinch(I + 1) = Temp
        End If
    Next I
Next Pass
For I = 1 To 23
    picresults.Print Names(I); Tab(50); Weight(I); "Lbs."
Next I
End Sub

Private Sub cmdweight_Click()
picresults.Cls
Open "N:\CS130\handin\Mol, Brandon\bikes.txt" For Input As #1
    For I = 1 To 23
        Input #1, Names(I)
        Input #1, Prices(I)
        Input #1, Weight(I)
        Input #1, Milespergallon(I)
        Input #1, Cubicinch(I)
    Next I
Close #1
    picresults.Print "Names"; Tab(40); "Weight(Lbs.)"
    picresults.Print "*****************************************************************************************************************************"
    For I = 1 To 23
        picresults.Print Names(I); Tab(40); Weight(I); "Lbs."
    Next I
cmdsortweight.Enabled = True
cmdfindweight.Enabled = True
End Sub

Private Sub Command1_Click()
picresults.Cls
cmdsortweight.Enabled = False
cmdfindweight.Enabled = False
Form1.Show
Form8.Hide
End Sub
