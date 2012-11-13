VERSION 5.00
Begin VB.Form frmTheatre 
   BackColor       =   &H00FF80FF&
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   12210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdbacktheatre 
      BackColor       =   &H0000FFFF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Width           =   1935
   End
   Begin VB.PictureBox picresults2 
      BackColor       =   &H0000FFFF&
      Height          =   3015
      Left            =   600
      ScaleHeight     =   2955
      ScaleWidth      =   10755
      TabIndex        =   2
      Top             =   2400
      Width           =   10815
   End
   Begin VB.CommandButton cmdload 
      BackColor       =   &H0000FFFF&
      Caption         =   "Load Show Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label LblTheatrestart 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Start by loading show options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "frmTheatre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Things to do in NYC
'Form Name: frmStart
'Author: Jake Johnson
'Date Written: 3/23/09
'Objective: Makes available the Theatre form and options


'Goes back to starting form
Private Sub cmdbacktheatre_Click()
frmTheatre.Hide
FrmStart.show
End Sub

'Loads Theatre options
Private Sub cmdload_Click()
Dim choice(1 To 5) As Integer, show(1 To 5) As String, mezzanine(1 To 5) As Integer, orchestra(1 To 5) As Integer, premium(1 To 5) As Integer, ctr As Integer

ctr = 0
Open App.Path & "\Theatre.txt" For Input As #1

Do While Not EOF(1)
    ctr = ctr + 1
    
    Input #1, choice(ctr), show(ctr), mezzanine(ctr), orchestra(ctr), premium(ctr)
Loop
Close

picresults2.Print "Option", "Show", Tab(20), "Mezzanine Cost", "Orchestra Cost", "Premium Cost"

For j = 1 To ctr
    picresults2.Print choice(j), show(j), Tab(20), FormatCurrency(mezzanine(j)), FormatCurrency(orchestra(j)), FormatCurrency(premium(j))
Next j
End Sub


