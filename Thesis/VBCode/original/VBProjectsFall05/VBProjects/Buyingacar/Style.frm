VERSION 5.00
Begin VB.Form StyleForm 
   BackColor       =   &H80000012&
   Caption         =   "Style"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMainForm1 
      Caption         =   "Back to Main Form 1"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8880
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox picSort 
      BackColor       =   &H80000013&
      Height          =   4815
      Left            =   360
      ScaleHeight     =   4755
      ScaleWidth      =   10035
      TabIndex        =   2
      Top             =   2520
      Width           =   10095
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00008000&
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
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblStyles 
      BackColor       =   &H0080FF80&
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
      Left            =   4080
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblStlye 
      BackColor       =   &H0080FF80&
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "StyleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Make As String
Dim J As Integer
Dim K As Integer
Dim L As Integer
Dim Pass As Integer
Dim I As Integer
Dim TempPrice As Double

Private Sub cmdMainForm1_Click()
StyleForm.Hide
MainForm1.Show ' Sends user back to MainForm1

End Sub

Private Sub cmdSort_Click()
 picSort.Print "Model", , , "Style", , , "Company", , , "Price"
    picSort.Print "******************************************************************************************************************************************************************************************"
    Dim MyStyle As String
    Dim CTR As Integer
    MyStyle = InputBox("Which style of car would you like to compare?")
    For CTR = 1 To 39
        If Style(CTR) = MyStyle Then
            picSort.Print Model(CTR), , Style(CTR), , Company(CTR), , FormatCurrency(Price(CTR)) 'Recieves specific style from user in InputBox and compares the specific style of all the companies
        End If
    Next CTR
    
    
End Sub
