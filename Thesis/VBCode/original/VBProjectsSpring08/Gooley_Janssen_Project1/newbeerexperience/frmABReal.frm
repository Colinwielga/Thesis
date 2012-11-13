VERSION 5.00
Begin VB.Form frmAB 
   BackColor       =   &H00004080&
   Caption         =   "Anheuser-Busch Form"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13890
   LinkTopic       =   "Form2"
   ScaleHeight     =   9435
   ScaleWidth      =   13890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000C0&
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   360
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   9960
      Picture         =   "frmABReal.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   2715
      TabIndex        =   8
      Top             =   7200
      Width           =   2775
   End
   Begin VB.CommandButton cmdNatural 
      BackColor       =   &H000000C0&
      Caption         =   "Natural Light"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   3855
   End
   Begin VB.CommandButton cmdBusch 
      BackColor       =   &H000000C0&
      Caption         =   "Busch"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   3855
   End
   Begin VB.CommandButton cmdMichelob 
      BackColor       =   &H000000C0&
      Caption         =   "Michelob"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
      Width           =   3855
   End
   Begin VB.CommandButton cmdBud 
      BackColor       =   &H000000C0&
      Caption         =   "Budweiser"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   3855
   End
   Begin VB.CommandButton cmdOrder 
      BackColor       =   &H000000C0&
      Caption         =   "Put the Beers in Chronological Order of Release"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   9135
      Left            =   4440
      ScaleHeight     =   9075
      ScaleWidth      =   8235
      TabIndex        =   1
      Top             =   240
      Width           =   8295
   End
   Begin VB.CommandButton cmdBrands 
      BackColor       =   &H000000C0&
      Caption         =   "(Click First)   List all of the Beers that Anheuser-Busch Offers"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblClicktoLearn 
      BackColor       =   &H000000C0&
      Caption         =   "Click on a Beer to Learn More About It"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   4800
      Width           =   3015
   End
End
Attribute VB_Name = "frmAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Beer Experience
'frm AB
'Lauren Gooley and Tim Janssen
'March 21, 2008
'This form is used to learn more about the Anheuser-Busch company and the beers it offers.

Option Explicit
Dim CTR As Single
Dim ABBeers(1 To 100) As String, Dates(1 To 100) As Integer

'This subroutine opens the array of all beers offered by the Anheuser-Busch company.
Private Sub cmdBrands_Click()
Dim J As Integer
Open App.Path & "\ABBeers.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, ABBeers(CTR), Dates(CTR)
Loop
For J = 1 To CTR
    picResults.Print ABBeers(J), ; Tab(50); Dates(J)
Next J
Close #1
End Sub

Private Sub cmdBud_Click()
MsgBox ("Brewed and sold since 1876, Budweiser leads the U.S. premium beer category, outselling all other domestic premium beers combined.")
End Sub

Private Sub cmdBusch_Click()
MsgBox ("Introduced in 1955, Busch has a smooth, light taste. The brand is the country's largest-selling subpremium-priced beer in all major demographics.")
End Sub

Private Sub cmdMichelob_Click()
MsgBox ("Michelob is a malty and full-bodied lager with an elegant European hop profile.")
End Sub

Private Sub cmdNatural_Click()
MsgBox ("Naturally brewed and less filling, low-calorie Natural Light was introduced in 1977.")
End Sub
'This subroutine sorts the types of beer in order of their chronilogical release.
Private Sub cmdOrder_Click()
Dim Pass As Single, POS As Single, J As Single, TempDates As Integer
Dim TempBrand As String
picResults.Cls
For Pass = 1 To CTR - 1
    For POS = 1 To CTR - Pass
        If Dates(POS) > Dates(POS + 1) Then
            TempDates = Dates(POS)
            Dates(POS) = Dates(POS + 1)
            Dates(POS + 1) = TempDates
            TempBrand = ABBeers(POS)
            ABBeers(POS) = ABBeers(POS + 1)
            ABBeers(POS + 1) = TempBrand
        End If
    Next POS
Next Pass
For J = 1 To CTR
    picResults.Print ABBeers(J), ; Tab(50); Dates(J)
Next J
End Sub

Private Sub cmdBack_Click()
frmAB.Hide
Companies.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub
