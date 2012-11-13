VERSION 5.00
Begin VB.Form frmDining 
   BackColor       =   &H00008000&
   Caption         =   "Dining"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   6120
      Picture         =   "Dining.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   4635
      TabIndex        =   6
      Top             =   1200
      Width           =   4695
   End
   Begin VB.CommandButton cmdGoBackHome2 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Caribbean Home Page"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   2535
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CommandButton cmdElegant 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Click to view elegant dining menu options"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdCasual 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Click to view casual dining menu options"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      Height          =   3375
      Left            =   240
      ScaleHeight     =   3315
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   3000
      Width           =   5655
   End
   Begin VB.Label lblDining 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Dining"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmDining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmDining
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form includes command buttons that shows the user either a casual dinner menu with prices
'included or an elegant dinner menu with prices included.

Private Sub cmdCasual_Click()
Dim CasualFood(1 To 30) As String, CasualPrice(1 To 100) As Single
Dim CTR As Integer
CTR = 0

picResults.Cls
Open App.Path & "\CasualDiningOptions.txt" For Input As #1

picResults.Print "Food Item"; Tab(35); "Price"
picResults.Print "***************************************************"

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, CasualFood(CTR), CasualPrice(CTR)
    picResults.Print CasualFood(CTR); Tab(35); FormatCurrency(CasualPrice(CTR), 2)
Loop
Close #1
End Sub

Private Sub cmdClear_Click()
picResults.Cls
End Sub

Private Sub cmdElegant_Click()
Dim ElegantFood(1 To 30) As String, ElegantPrice(1 To 100) As Single
Dim CTR As Integer
CTR = 0

picResults.Cls
Open App.Path & "\ElegantDiningOptions.txt" For Input As #1

picResults.Print "Food Item"; Tab(35); "Price"
picResults.Print "***************************************************"

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, ElegantFood(CTR), ElegantPrice(CTR)
    picResults.Print ElegantFood(CTR); Tab(35); FormatCurrency(ElegantPrice(CTR), 2)
Loop
Close #1

End Sub

Private Sub cmdGoBackHome2_Click()
frmCaribbeanHome.Show
frmDining.Hide
End Sub

