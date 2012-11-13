VERSION 5.00
Begin VB.Form EndOfDay 
   BackColor       =   &H00800080&
   Caption         =   "End of Day Totals"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Main"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00FF0000&
      Caption         =   "Show Totals"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   1200
      ScaleHeight     =   6555
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "EndOfDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMain_Click()
MovieMain.Show
EndOfDay.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdTotal_Click()
    'Initialize counter
CTR = 0
    'open file
Open "M:\cs130\Stevens_Jackie\MovieFile.txt" For Input As #1
    'load array
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Screen(CTR), Movie(CTR), Rating(CTR), Time1(CTR), Time2(CTR), Time3(CTR), Time4(CTR), Time5(CTR)
Loop
Close
    'Print totals per movie
picResults.Print Movie(1); "="; Tab(70); FormatCurrency(MovieTotal1)
picResults.Print Movie(2); "="; Tab(70); FormatCurrency(MovieTotal2)
picResults.Print Movie(3); "="; Tab(70); FormatCurrency(MovieTotal3)
picResults.Print Movie(4); "="; Tab(70); FormatCurrency(MovieTotal4)
picResults.Print Movie(5); "="; Tab(70); FormatCurrency(MovieTotal5)
picResults.Print Movie(6); "="; Tab(70); FormatCurrency(MovieTotal6)
picResults.Print Movie(7); "="; Tab(70); FormatCurrency(MovieTotal7)
picResults.Print Movie(8); "="; Tab(70); FormatCurrency(MovieTotal8)
picResults.Print Movie(9); "="; Tab(70); FormatCurrency(MovieTotal9)
picResults.Print Movie(10); "="; Tab(70); FormatCurrency(MovieTotal10)
picResults.Print Movie(11); "="; Tab(70); FormatCurrency(MovieTotal11)
picResults.Print Movie(12); "="; Tab(70); FormatCurrency(MovieTotal12)
picResults.Print Movie(13); "="; Tab(70); FormatCurrency(MovieTotal13)
picResults.Print Movie(14); "="; Tab(70); FormatCurrency(MovieTotal14)
picResults.Print Movie(15); "="; Tab(70); FormatCurrency(MovieTotal15)
picResults.Print Movie(16); "="; Tab(70); FormatCurrency(MovieTotal16)
picResults.Print Movie(17); "="; Tab(70); FormatCurrency(MovieTotal17)
picResults.Print Movie(18); "="; Tab(70); FormatCurrency(MovieTotal2)
    'Print total for the end of the day
picResults.Print "-------------------------------------------------------------------------------------------------------------------------------------------------"
picResults.Print "End Of Day Total ="; Tab(70); FormatCurrency(EndTotal)
End Sub
