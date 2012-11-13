VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00808000&
   Caption         =   "Engine Sizes"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form10"
   ScaleHeight     =   8640
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfindsize 
      Caption         =   "Click to find the  engine size closest to your preference and their corresponding models"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   480
      TabIndex        =   4
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton cmdsortsize 
      Caption         =   "Click to list the engine size from largest to smallest"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   480
      TabIndex        =   3
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click to return to the main menu"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton cmdlistsize 
      Caption         =   "Click to see the size of the engines of all models"
      Height          =   1935
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H8000000E&
      Height          =   6975
      Left            =   4440
      ScaleHeight     =   6915
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   360
      Width           =   7335
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Names(1 To 23) As String
Dim Prices(1 To 23) As Integer
Dim Weight(1 To 23) As Single
Dim Milespergallon(1 To 23) As Integer
Dim Cubicinch(1 To 23) As Single
Dim Enginesizes(1 To 4) As Single
Dim I As Integer
Dim x As Integer

Private Sub cmdfindsize_Click()
picresults.Cls
usersinput = InputBox("enter an engine size you prefer")
Open "N:\CS130\handin\Mol, Brandon\engineweights.txt" For Input As #1
    For x = 1 To 4
        Input #1, Enginesizes(x)
    Next x
Close #1
For Pass = 1 To 4
    For x = 1 To 4 - Pass
        If Abs(Enginesizes(x) - usersinput) >= Abs(Enginesizes(x + 1) - usersinput) Then
        Temp = Enginesizes(x)
        Enginesizes(x) = Enginesizes(x + 1)
        Enginesizes(x + 1) = Temp
        End If
    Next x
Next Pass
picresults.Print "The engine size(Cu. In.) closest to your preference is"
picresults.Print "*************************************************************************************************************************************************************************************************"
picresults.Print Enginesizes(1); "In.^3"
picresults.Print
picresults.Print
picresults.Print "Names of the models corresponding to that size"; Tab(60); "Engine sizes of those models"
picresults.Print "***************************************************************************************************************************************8**********************************************************************************"
For I = 1 To 23
    If Enginesizes(1) = Cubicinch(I) Then
    picresults.Print Names(I); Tab(60); Cubicinch(I); "In.^3"
    End If
Next I
End Sub

Private Sub cmdlistsize_Click()
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
    picresults.Print "Names"; Tab(40); "Engine sizes(Cu. In.)"
    picresults.Print "*********************************************************************************************************************************************************************************************"
    For I = 1 To 23
        picresults.Print Names(I); Tab(40); Cubicinch(I); "In.^3"
    Next I
cmdsortsize.Enabled = True
cmdfindsize.Enabled = True
End Sub

Private Sub cmdsortsize_Click()
picresults.Cls
picresults.Print "Names"; Tab(40); "Engine size fron largest to smallest( In Cu. In.)"
picresults.Print "************************************************************************************************************************************************************************************************"
For Pass = 1 To 22
    For I = 1 To 23 - Pass
        If Cubicinch(I) <= Cubicinch(I + 1) Then
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
    picresults.Print Names(I); Tab(50); Cubicinch(I); "In.^3"
Next I
End Sub

Private Sub Command1_Click()
picresults.Cls
cmdsortsize.Enabled = False
cmdfindsize.Enabled = False
Form1.Show
Form10.Hide
End Sub
