VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00000080&
   Caption         =   "Miles Per Gallon"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form9"
   ScaleHeight     =   8685
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfindmpg 
      Caption         =   "Click to list the models with an acceptable miles per gallon"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   360
      TabIndex        =   4
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdsortmpg 
      Caption         =   "Click to list the miles per gallon from highest to lowest"
      Enabled         =   0   'False
      Height          =   1815
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click to return to the main menu"
      Height          =   1695
      Left            =   480
      TabIndex        =   2
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton cmdlist 
      Caption         =   "Click to see the miles per gallon of all models"
      Height          =   1815
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H8000000E&
      Height          =   6015
      Left            =   4440
      ScaleHeight     =   5955
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   360
      Width           =   6255
   End
End
Attribute VB_Name = "Form9"
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
Dim Found As Boolean
Dim CTR As Integer
Private Sub cmdfindmpg_Click()
usersinput = InputBox("enter a miles per gallon that is acceptable")
picresults.Cls
picresults.Print "Names"; Tab(40); "Miles per gallon of each acceptable model"
picresults.Print "***************************************************************************************************************************************************************************************"
Found = False
CTR = 0
For I = 1 To 23
    If Milespergallon(I) >= usersinput Then
    picresults.Print Names(I); Tab(50); Milespergallon(I); "Mpg"
    Else
    CTR = CTR + 1
    End If
Next I
    If CTR = 23 Then
    Found = True
    Else
    Found = False
    End If
    If Found = True Then
    MsgBox ("There are no models with an acceptable miles per gallon for you")
    End If
End Sub
Private Sub cmdlist_Click()
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
    picresults.Print "Names"; Tab(40); "Miles per gallon"
    picresults.Print "***************************************************************************************************************************************"
    For I = 1 To 23
        picresults.Print Names(I); Tab(40); Milespergallon(I); "Mpg"
    Next I
cmdsortmpg.Enabled = True
cmdfindmpg.Enabled = True
End Sub

Private Sub cmdsortmpg_Click()
picresults.Cls
picresults.Print "Names"; Tab(40); "Miles per gallon from highest to lowest"
picresults.Print "********************************************************************************************************************************************************************"
For Pass = 1 To 22
    For I = 1 To 23 - Pass
        If Milespergallon(I) <= Milespergallon(I + 1) Then
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
    picresults.Print Names(I); Tab(50); Milespergallon(I); "Mpg"
Next I
End Sub

Private Sub Command1_Click()
picresults.Cls
cmdsortmpg.Enabled = False
cmdfindmpg.Enabled = False
Form1.Show
Form9.Hide
End Sub
