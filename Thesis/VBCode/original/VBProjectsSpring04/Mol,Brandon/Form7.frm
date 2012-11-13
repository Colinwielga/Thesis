VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00808080&
   Caption         =   "Prices"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form7"
   ScaleHeight     =   8490
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdlocateprice 
      Caption         =   "List the models in your price range"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   360
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdfindprice 
      Caption         =   "List the models in order from highest price to smallest price"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click to return to the main menu"
      Height          =   1815
      Left            =   360
      TabIndex        =   2
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdprice 
      Caption         =   "Click to see prices of all the models"
      Height          =   1695
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H8000000E&
      Height          =   6015
      Left            =   3240
      ScaleHeight     =   5955
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   480
      Width           =   6255
   End
End
Attribute VB_Name = "Form7"
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
Private Sub cmdfindprice_Click()
picresults.Cls
picresults.Print "Names"; Tab(40); "Prices from highest to smallest"
picresults.Print "******************************************************************************************************************************************"
For Pass = 1 To 22
    For I = 1 To 23 - Pass
        If Prices(I) <= Prices(I + 1) Then
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
    picresults.Print Names(I); Tab(50); FormatCurrency(Prices(I))
Next I

End Sub

Private Sub cmdlocateprice_Click()
usersinput = InputBox("enter a price in your range")
picresults.Cls
picresults.Print "Names"; Tab(40); "Prices of models in your price range"
picresults.Print "***************************************************************************************************************************************************************"
CTR = 0
For I = 1 To 23
    If Prices(I) <= usersinput Then
    picresults.Print Names(I); Tab(50); FormatCurrency(Prices(I))
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
    MsgBox ("There are no models in your price range")
    End If
End Sub

Private Sub cmdprice_Click()
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
    picresults.Print "Names"; Tab(40); "Prices"
    picresults.Print "************************************************************************************************************************"
    For I = 1 To 23
        picresults.Print Names(I); Tab(40); FormatCurrency(Prices(I))
    Next I
cmdfindprice.Enabled = True
cmdlocateprice.Enabled = True
End Sub

Private Sub Command1_Click()
picresults.Cls
cmdfindprice.Enabled = False
cmdlocateprice.Enabled = False
Form1.Show
Form7.Hide
End Sub



