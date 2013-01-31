VERSION 5.00
Begin VB.Form frmStartPage
   BackColor       =   &H00000000&
   Caption         =   "Start Page"
   ClientHeight    =   12705
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15990
   LinkTopic       =   "Form2"
   ScaleHeight     =   12705
   ScaleWidth      =   15990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAppCar
      BackColor       =   &H00808080&
      Caption         =   "Find Appropriate Car For You"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9600
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   11400
      Width           =   1575
   End
   Begin VB.CommandButton cmdFindCar
      BackColor       =   &H00808080&
      Caption         =   "Find  Desired Car For You "
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9600
      Width           =   2775
   End
   Begin VB.CommandButton cmdMercedes
      BackColor       =   &H80000008&
      Caption         =   "Mercedes"
      Height          =   3015
      Left            =   6600
      Picture         =   "frmStartPage.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8280
      Width           =   2775
   End
   Begin VB.CommandButton cmdVolvo
      BackColor       =   &H80000008&
      Caption         =   "Volvo"
      Height          =   3015
      Left            =   12840
      Picture         =   "frmStartPage.frx":0BC0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8880
      Width           =   2775
   End
   Begin VB.CommandButton cmdSeat
      BackColor       =   &H80000008&
      Caption         =   "Seat"
      Height          =   3015
      Left            =   360
      Picture         =   "frmStartPage.frx":1666
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8880
      Width           =   2775
   End
   Begin VB.CommandButton cmdAudi
      Caption         =   "Audi"
      Height          =   3015
      Left            =   12840
      Picture         =   "frmStartPage.frx":1DC3
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   2775
   End
   Begin VB.CommandButton cmdAlfaRomeo
      BackColor       =   &H80000008&
      Caption         =   "Alfa Romeo"
      Height          =   3015
      Left            =   12840
      Picture         =   "frmStartPage.frx":429F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   2775
   End
   Begin VB.CommandButton cmdPeugeot
      BackColor       =   &H80000008&
      Caption         =   "Peugeot"
      Height          =   3015
      Left            =   360
      Picture         =   "frmStartPage.frx":58CC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   2775
   End
   Begin VB.CommandButton cmdBMW
      BackColor       =   &H80000007&
      Caption         =   "BMW"
      Height          =   3015
      Left            =   360
      MaskColor       =   &H00000000&
      Picture         =   "frmStartPage.frx":738D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   2775
   End
   Begin VB.PictureBox Picture1
      Height          =   5295
      Left            =   4440
      Picture         =   "frmStartPage.frx":8057
      ScaleHeight     =   5235
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   2880
      Width           =   7095
   End
   Begin VB.Label CrBrands
      BackColor       =   &H00000000&
      Caption         =   "Chose the Automaker"
      BeginProperty Font
         Name            =   "Bernard MT Condensed"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   12015
   End
End
Attribute VB_Name = "frmStartPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'hide all forms to keep the screen less cluttered
Private Sub cmdAlfaRomeo_Click()
    'hide Start page from user
    frmStartPage.Hide
    'show Alfa Romeo page to user
    frmAlfa.Show
End Sub

Private Sub cmdVolvo_Click()
    'hide Start page from user
    frmStartPage.Hide
    'show Volvo page to user
    frmVolvo.Show
End Sub

'This program will compare inputed data with given age ragange and gender and show appropriate output
Private Sub cmdAppCar_Click()
Dim Gender As String, Age As Single
    'asignes gender from input box as Gender to search within given option F of M
    Gender = InputBox(" Enter your Gender, F for Female , M for Male.", "Gender")
    'asignes age  from input box as Age to search within  given range
    Age = InputBox(" Enter your Age.", "Age")

       If (Gender = "M" Or Gender = "F") And Age >= 5 And Age < 18 Then
                'hide Start page from user
                frmStartPage.Hide
                'show Bicycle page to user
                frmBicycle.Show
        ElseIf Gender = "M" And Age > 24 And 35 > Age Then
                'show Old Fiat page to user
                frmOldFiat.Show
                'hide Start page from user
                frmStartPage.Hide
        ElseIf Gender = "M" And 18 < Age And Age <= 21 Then
                'hide Start page from user
                frmStartPage.Hide
                'show Trabant page to user
                frmTrabant.Show
        ElseIf Gender = "M" And Age > 21 And Age <= 24 Then
                'hide Start page from user
                frmStartPage.Hide
                'show TrabantS page to user
                frmTrabantS.Show
        ElseIf "F" = Gender And Age > 24 And Age <= 35 Then
                'show Beetle Concept page to user
                frmBeetleC.Show
                'hide Start page from user
                frmStartPage.Hide
        ElseIf ("F" = Gender Or Gender = "M") And Age > 35 And 70 > Age Then
                'hide Start page from user
                frmStartPage.Hide
                'show Beetle page to user
                frmBeetle.Show
        ElseIf Gender = "F" And 18 < Age And Age <= 21 Then
                'hide Start page from user
                frmStartPage.Hide
                'show Fiat 500 page to user
                frmNewFiat.Show
        ElseIf Gender = "F" And Age > 21 And 24 > Age Then
                'show Mini page to user
                frmMini.Show
                'hide Start page from user
                frmStartPage.Hide
        Else
                'messagebox to user should not drive
                MsgBox "" & UserName & ", You Should Not Drive", , "Warning."

        End If
End Sub

Private Sub cmdAudi_Click()
    'show Audi page to user
    frmAudi.Show
    'hide Start page from user
    frmStartPage.Hide
End Sub

Private Sub cmdBMW_Click()
    'hide Start page from user
    frmStartPage.Hide
    'shows BMW page to user
    frmBMW.Show
End Sub

Private Sub cmdQuit_Click()
    'messagebox to user, thanks for wisiting
    MsgBox "" & UserName & ", Thanks For Visiting Us", , "Have a Safe Trip."
    'quit program
    End
End Sub

Private Sub cmdFindCar_Click()
    'hide Start page from user
    frmStartPage.Hide
    'show Find Car page to user
    frmFindCar.Show
End Sub

Private Sub cmdMercedes_Click()
    'hide Start page from user
    frmStartPage.Hide
    'show Mercedes page to user
    frmMercedes.Show
End Sub

Private Sub cmdPeugeot_Click()
    'show Peugeot page to user
    frmPeugeot.Show
    'hide Start page from user
    frmStartPage.Hide
End Sub

Private Sub cmdSeat_Click()
    'hide Start page from user
    frmStartPage.Hide
    'show Seat page to user
    frmSeat.Show
End Sub

