VERSION 5.00
Begin VB.Form frmTwo 
   BackColor       =   &H80000003&
   Caption         =   "Compare cars"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4455
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H8000000D&
      Caption         =   "Back to the Shop"
      Height          =   735
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdMPG 
      BackColor       =   &H8000000D&
      Caption         =   "By MPG"
      Height          =   735
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrice 
      BackColor       =   &H8000000D&
      Caption         =   "By Price"
      Height          =   735
      Left            =   240
      MaskColor       =   &H00800000&
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H80000009&
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Image img2 
      Height          =   1155
      Left            =   1320
      Picture         =   "frmTwo.frx":0000
      Top             =   3960
      Width           =   1740
   End
   Begin VB.Label lblCompare 
      BackColor       =   &H80000003&
      Caption         =   "COMPARE CARS BY:"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pos, Size, Pass As Integer

Private Sub cmdBack_Click()
    frmOne.Visible = True
    frmTwo.Visible = False
End Sub

Private Sub cmdMPG_Click()
    Dim TempPrice, TempMPG As Single
    Dim TempCar As String
    
    For Pass = 1 To (Size - 1)
        For pos = 1 To (Size - Pass)
            If MPG(pos) > MPG(pos + 1) Then
                TempMPG = MPG(pos)
                MPG(pos) = MPG(pos + 1)
                MPG(pos + 1) = TempMPG
            
                TempCar = Car(pos)
                Car(pos) = Car(pos + 1)
                Car(pos + 1) = TempCar
                
                TempPrice = Price(pos)
                Price(pos) = Price(pos + 1)
                Price(pos + 1) = TempPrice
            End If
        Next pos
    Next Pass
    picResults2.Cls
    picResults2.Print "Car name:", , "Mileage:"
    picResults2.Print "***********************************************"
        For pos = 1 To Size
        picResults2.Print Car(pos), , MPG(pos)
        Next pos
End Sub

Private Sub cmdPrice_Click()

    Dim TempPrice, TempMPG As Single
    Dim TempCar As String
    
    For Pass = 1 To (Size - 1)
        For pos = 1 To (Size - Pass)
            If Price(pos) > Price(pos + 1) Then
                TempPrice = Price(pos)
                Price(pos) = Price(pos + 1)
                Price(pos + 1) = TempPrice
            
                TempCar = Car(pos)
                Car(pos) = Car(pos + 1)
                Car(pos + 1) = TempCar
                
                TempMPG = MPG(pos)
                MPG(pos) = MPG(pos + 1)
                MPG(pos + 1) = TempMPG
            End If
        Next pos
    Next Pass
        picResults2.Cls
        picResults2.Print "Car name:", , "Base Price:"
        picResults2.Print "***********************************************"
        For pos = 1 To Size
        picResults2.Print Car(pos), , FormatCurrency(Price(pos))
        Next pos
        
End Sub

Private Sub Form_Load()
 
    pos = 0
    Open App.Path & "\CarData.txt" For Input As #2
    Do Until EOF(2)
        pos = pos + 1
        Input #2, Car(pos), Price(pos), MPG(pos)
    Loop
    Close #2
    Size = pos
    
End Sub


