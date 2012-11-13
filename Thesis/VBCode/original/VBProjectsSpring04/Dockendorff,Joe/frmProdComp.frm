VERSION 5.00
Begin VB.Form frmProdComp 
   BackColor       =   &H000000FF&
   Caption         =   "Computers"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   9720
      TabIndex        =   14
      Top             =   7920
      Width           =   615
   End
   Begin VB.CommandButton cmdNext 
      Cancel          =   -1  'True
      Caption         =   "Go to next"
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   7920
      Width           =   1695
   End
   Begin VB.OptionButton optComP 
      BackColor       =   &H000000FF&
      Caption         =   "Compaq Presario 2200+ Desktop with 17-Inch Flat-Screen Monitor"
      Height          =   615
      Left            =   7680
      TabIndex        =   12
      Top             =   4320
      Width           =   2535
   End
   Begin VB.OptionButton optHPFlatP 
      BackColor       =   &H000000FF&
      Caption         =   "HP Pavilion 3.0E GHz Desktop with 15-Inch Flat-Panel Monitor and Color Printer"
      Height          =   855
      Left            =   4320
      TabIndex        =   11
      Top             =   4200
      Width           =   2295
   End
   Begin VB.OptionButton optEmach 
      BackColor       =   &H000000FF&
      Caption         =   "eMachines 2.7GHz Desktop with 17-Inch Flat-Screen Monitor and Lexmark Printer"
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   4320
      Width           =   2535
   End
   Begin VB.OptionButton optSony 
      BackColor       =   &H000000FF&
      Caption         =   "Sony VAIO 2.8C GHz Desktop with 15-Inch Flat-Panel Monitor and Lexmark Printer"
      Height          =   735
      Left            =   7680
      TabIndex        =   9
      Top             =   720
      Width           =   2655
   End
   Begin VB.OptionButton optComX 
      BackColor       =   &H000000FF&
      Caption         =   "Compaq X Gaming Desktop with Intel® Pentium® 4 Processor 3.0GHz"
      Height          =   735
      Left            =   4320
      TabIndex        =   8
      Top             =   720
      Width           =   2295
   End
   Begin VB.OptionButton optHPFlatS 
      BackColor       =   &H000000FF&
      Caption         =   "HP Pavilion 2.7GHz Desktop with 17-Inch Flat-Screen Monitor and Color Printer"
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   840
      Width           =   2655
   End
   Begin VB.PictureBox Picture6 
      Height          =   2175
      Left            =   4680
      Picture         =   "frmProdComp.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   2535
      Left            =   7320
      Picture         =   "frmProdComp.frx":0F20
      ScaleHeight     =   2475
      ScaleWidth      =   3315
      TabIndex        =   5
      Top             =   5160
      Width           =   3375
   End
   Begin VB.PictureBox Picture4 
      Height          =   2535
      Left            =   120
      Picture         =   "frmProdComp.frx":426B
      ScaleHeight     =   2475
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   5160
      Width           =   3375
   End
   Begin VB.PictureBox Picture3 
      Height          =   2535
      Left            =   7440
      Picture         =   "frmProdComp.frx":6D83
      ScaleHeight     =   2475
      ScaleWidth      =   3195
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   3720
      Picture         =   "frmProdComp.frx":9D58
      ScaleHeight     =   2475
      ScaleWidth      =   3315
      TabIndex        =   2
      Top             =   5160
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   120
      Picture         =   "frmProdComp.frx":CEDD
      ScaleHeight     =   2475
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label labelComp 
      BackColor       =   &H000000FF&
      Caption         =   "Choose a computer and click Go to next:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frmProdComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjElectrPlus (Joe Dockendorff's VB Project.vbp)
'Form Name : frmProdComp (frmProdComp.frm)
'Author: Joe Dockendorff
'Date Written: March 13, 2004
'Purpose of Form: This form provides the user with six different
                 'desktop computers.  The user chooses an option
                 'and the "go to next" button becomes available.
                 'The options are each assigned a number in their
                 'array and is carried over to the subtitle page.
                 
'Option Explicit is a command to force
'the user to declare all variables
'before they can be used.
Option Explicit

Private Sub cmdNext_Click()
'This button figures out which option is used and assigns it a number.
'Also, the ProdComp form is hidden and the Start Form is shown

If optHPFlatS = True Then
    C = 1
ElseIf optSony = True Then
    C = 2
ElseIf optHPFlatP = True Then
    C = 3
ElseIf optEmach = True Then
    C = 4
ElseIf optComP = True Then
    C = 5
ElseIf optComX = True Then
    C = 6
End If

frmProdComp.Hide
frmStart.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
'This opens the necessary information and loads it into arrays, disables the Next button
'until the user picks a computer.

ReDim Comp(1 To 6) As String
ReDim CompPrice(1 To 6) As Single
Path = "N:\CS130\handin\Dockendorff, Joe\"

'Open the file associated with the product, in this case, the file
'containing computer information.
Close #1
Open Path & "comp.txt" For Input As #1

For C = 1 To 6
    Input #1, Comp(C), CompPrice(C)
Next C
Close #1
cmdNext.Enabled = False
End Sub

Private Sub optComP_Click()
cmdNext.Enabled = True
End Sub

Private Sub optComX_Click()
cmdNext.Enabled = True
End Sub

Private Sub optEmach_Click()
cmdNext.Enabled = True
End Sub

Private Sub optHPFlatP_Click()
cmdNext.Enabled = True
End Sub

Private Sub optHPFlatS_Click()
cmdNext.Enabled = True
End Sub

Private Sub optSony_Click()
cmdNext.Enabled = True
End Sub
