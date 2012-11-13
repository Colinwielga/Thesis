VERSION 5.00
Begin VB.Form Style 
   Caption         =   "Style"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSort 
      Height          =   3615
      Left            =   3240
      ScaleHeight     =   3555
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   3000
      Width           =   4095
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Compare this style within other companies"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   960
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtInput 
      Height          =   1215
      Left            =   3600
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label lblStyles 
      Caption         =   "Truck, Sedan, SUV, Converrtible"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblStlye 
      Caption         =   "Which style of car are you interested in?"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "Style"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Style(1 To 39) As Integer
Dim Price(1 To 39) As Double
Dim J As Integer
Dim K As Integer
Dim L As Integer
Dim Pass As Integer







Private Sub cmdSort_Click()
Style = InputBox("Which style of car would you like to compare?")
For Pass = 1 To J - 1
    For K = 1 To J - Pass
        If Price(K) > Price(K + 1) Then
           Name = Price(K)
           Price(K) = Price(K + 1)
           Price(K + 1) = Name
        End If
    Next K
Next Pass
picSort.Print "Name", "Company", "Price"
picSort.Print "************************"
For L = 1 To J
    picSort.Print Name(L), Company(L), Price(L)

Next L

End Sub
