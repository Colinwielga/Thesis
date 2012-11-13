VERSION 5.00
Begin VB.Form frmFirstArbor 
   BackColor       =   &H00808000&
   Caption         =   "First Electric Arbor"
   ClientHeight    =   12120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   ScaleHeight     =   12120
   ScaleWidth      =   13020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMoveOn 
      Caption         =   "To the Second Electric!"
      Height          =   495
      Left            =   9120
      TabIndex        =   26
      Top             =   10320
      Width           =   2655
   End
   Begin VB.CommandButton cmdTotal2 
      Caption         =   "Calculate Arbor Weight"
      Height          =   495
      Left            =   9120
      TabIndex        =   25
      Top             =   9600
      Width           =   2655
   End
   Begin VB.PictureBox picTArbor 
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   24
      Top             =   10800
      Width           =   1335
   End
   Begin VB.PictureBox picTHang 
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   23
      Top             =   10800
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Quit"
      Height          =   495
      Left            =   9600
      TabIndex        =   22
      Top             =   11400
      Width           =   1695
   End
   Begin VB.CommandButton cmdTotal1 
      Caption         =   "Calculate Electric Weight"
      Height          =   495
      Left            =   9120
      TabIndex        =   21
      Top             =   8880
      Width           =   2655
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Show First Electric Information"
      Height          =   495
      Left            =   9120
      TabIndex        =   20
      Top             =   8160
      Width           =   2655
   End
   Begin VB.TextBox txtHeavy 
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   9600
      Width           =   1335
   End
   Begin VB.TextBox txtMedium 
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   9000
      Width           =   1335
   End
   Begin VB.TextBox txtLight 
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   8400
      Width           =   1335
   End
   Begin VB.TextBox txtS 
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   7
      Top             =   10200
      Width           =   1335
   End
   Begin VB.TextBox txtF 
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   9600
      Width           =   1335
   End
   Begin VB.TextBox txtP 
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   9000
      Width           =   1335
   End
   Begin VB.TextBox txtE 
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   3900
      Left            =   2880
      Picture         =   "frm5FirstArbor.frx":0000
      Top             =   720
      Width           =   6930
   End
   Begin VB.Label lbl9 
      BackStyle       =   0  'Transparent
      Caption         =   "Arbor weight = "
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   19
      Top             =   10920
      Width           =   1935
   End
   Begin VB.Label lbl8 
      BackStyle       =   0  'Transparent
      Caption         =   "35 lb bricks"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   18
      Top             =   9600
      Width           =   1935
   End
   Begin VB.Label lbl7 
      BackStyle       =   0  'Transparent
      Caption         =   "25 lb bricks"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   17
      Top             =   9000
      Width           =   1935
   End
   Begin VB.Label lbl6 
      BackStyle       =   0  'Transparent
      Caption         =   "12.5 lb bricks"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Top             =   8400
      Width           =   1935
   End
   Begin VB.Label lbl5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total hang weight = "
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   10920
      Width           =   2055
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "Scoops (5.5 lbs)"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   10200
      Width           =   1575
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fresnels (9 lbs)"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "PARs (7.5 lbs)"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ellipsoidals (15 lbs)"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   8400
      Width           =   2055
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm5FirstArbor.frx":D9F8
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      TabIndex        =   3
      Top             =   6600
      Width           =   11775
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm5FirstArbor.frx":DBD6
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1200
      TabIndex        =   2
      Top             =   5160
      Width           =   10455
   End
   Begin VB.Label lblArbor 
      BackStyle       =   0  'Transparent
      Caption         =   "The arbor of the counter-weight system in Gorecki/Escher"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   4680
      Width           =   4695
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "The Counter-Weight System - First Electric"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmFirstArbor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form deals with equalizing the weight values on the electric and in the arbor of the counter-weight system.
    Dim Ellipsoidal As Integer
    Dim PAR As Integer
    Dim Fresnel As Integer
    Dim Scoop As Integer
    Dim THang As Single
    
    Dim LWeight As Integer
    Dim MWeight As Integer
    Dim HWeight As Integer
    Dim TArbor As Single

Private Sub cmdDisplay_Click()
'Shows the inputted data of the first electric.
    frmFInfo.Show
    
End Sub

Private Sub cmdTotal1_Click()
'Totals the weight of the lights hung by multiplying the number of each by their respective weights.
    
    picTHang.Cls
    
    Ellipsoidal = txtE.Text
    PAR = txtP.Text
    Fresnel = txtF.Text
    Scoop = txtS.Text
    
    THang = (Ellipsoidal * 15) + (PAR * 7.5) + (Fresnel * 9) + (Scoop * 5.5)
    
    picTHang.Print THang
    
End Sub

Private Sub cmdTotal2_Click()
'Totals the weight added to the arbor by multiplying the number of bricks added by their respective weights.

    picTArbor.Cls
    
    LWeight = txtLight.Text
    MWeight = txtMedium.Text
    HWeight = txtHeavy.Text
    
    TArbor = (LWeight * 12.5) + (MWeight * 25) + (HWeight * 35)
    
    picTArbor.Print TArbor

    If TArbor > (THang + 7) Then
        MsgBox "At " & TArbor & " pounds, the counter-weight system is arbor-heavy. Take some weight off the arbor and try again!", , "Oops!"
    ElseIf TArbor < (THang - 7) Then
        MsgBox "At " & TArbor & " pounds, the counter-weight system is electric-heavy. Add some weight to the arbor and try again!", , "Oops!"
    Else
        MsgBox "At " & TArbor & " pounds, the arbor is in weight! Nice job! Let's move on to hanging the second electric!", , "Success!"
    End If
       
End Sub

Private Sub cmdMoveOn_Click()
'Advances the program to the Second Electric form.
    frmFirstArbor.Hide
    frmSecondE.Show

End Sub

Private Sub cmdEnd_Click()
'Ends the program
    End
End Sub

