VERSION 5.00
Begin VB.Form frmAdidas 
   BackColor       =   &H80000012&
   Caption         =   "BRAND : ADIDAS"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "exit"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   3135
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "back"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton cmdPrice 
      Caption         =   "sort by price"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton cmdShoes 
      Caption         =   "shoes"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   3
      Top             =   3240
      Width           =   3855
   End
   Begin VB.PictureBox picOutput 
      Height          =   2895
      Left            =   3480
      ScaleHeight     =   2835
      ScaleWidth      =   3795
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   240
      Picture         =   "frmAdidas.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by: Justin Swierk 2oo5"
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   3135
   End
End
Attribute VB_Name = "frmAdidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim shoeName(1 To 10) As String
Dim price(1 To 10) As Double
Dim quality(1 To 10) As Double
Dim I As Double, J As Double, K As Double
Dim PriceRange As Double, Pass As Double
Dim Temp As Double, Z As Double
Dim Temp2 As String, Temp3 As Double
Dim C As Double, L As Double











Private Sub cmdBack_Click(Index As Integer)
    frmAdidas.Hide
    frmSTARTUP.Show
End Sub

Private Sub cmdExit_Click(Index As Integer)
    End
End Sub

Private Sub cmdPrice_Click()
    picOutput.Cls
    
    For Pass = 1 To L - 1
        For Z = 1 To L - Pass
            If price(Z) > price(Z + 1) Then
                Temp = price(Z)
                price(Z) = price(Z + 1)
                price(Z + 1) = Temp
                Temp2 = shoeName(Z)
                shoeName(Z) = shoeName(Z + 1)
                shoeName(Z + 1) = Temp2
                Temp3 = quality(Z)
                quality(Z) = quality(Z + 1)
                quality(Z + 1) = Temp3
            
            
            
            End If
        Next Z
    Next Pass
    
    
    picOutput.Print "Name of Shoe",
    picOutput.Print "Price",
    picOutput.Print "Quality"
    picOutput.Print , , ,
    picOutput.Print "(5 = best)"
    picOutput.Print
    For J = 1 To I
        If PriceRange >= price(J) Then
            picOutput.Print shoeName(J),
            picOutput.Print FormatCurrency(price(J)),
            picOutput.Print quality(J)
            picOutput.Print
        End If
    Next J
    
End Sub

Private Sub cmdShoes_Click()
    picOutput.Cls
    
    cmdPrice.Visible = True
      
    Open App.Path & "\adidasShoes.txt" For Input As #9
    
    Do Until EOF(9)
        I = I + 1
        Input #9, shoeName(I), price(I), quality(I)
    Loop
    Close 9
        C = I
        L = I
    PriceRange = InputBox("HOW MUCH MONEY CAN YOU SPEND??", "MONEY AVAILABLE")
        If PriceRange < 54.89 Then
            picOutput.Print "SORRY, WE HAVE NO SHOE TO FIT YOUR PRICE RANGE"
        End If
                  
    picOutput.Print "Name of Shoe",
    picOutput.Print "Price",
    picOutput.Print "Quality"
    picOutput.Print , , ,
    picOutput.Print "(5 = best)"
    picOutput.Print
    For J = 1 To I
        If PriceRange >= price(J) Then
            picOutput.Print shoeName(J),
            picOutput.Print FormatCurrency(price(J)),
            picOutput.Print quality(J)
            picOutput.Print
        End If
    Next J
        
        
End Sub


