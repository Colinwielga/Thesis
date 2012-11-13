VERSION 5.00
Begin VB.Form Purchasepage
   Caption         =   "Customize your purchase here"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2
      Caption         =   "Exit"
      Height          =   495
      Left            =   5160
      TabIndex        =   22
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdcalculate
      Caption         =   "Calculate"
      Height          =   495
      Left            =   1560
      TabIndex        =   21
      Top             =   4920
      Width           =   2175
   End
   Begin VB.PictureBox Total
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   20
      Top             =   3840
      Width           =   1695
   End
   Begin VB.PictureBox Salestax
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   18
      Top             =   3240
      Width           =   1695
   End
   Begin VB.PictureBox subtotal
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   17
      Top             =   2640
      Width           =   1695
   End
   Begin VB.PictureBox AccessoriesFinish
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   16
      Top             =   2040
      Width           =   1695
   End
   Begin VB.PictureBox carprice
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   15
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox carmodeltext
      Height          =   495
      Left            =   6120
      TabIndex        =   14
      Top             =   840
      Width           =   1695
   End
   Begin VB.Frame Frame2
      Caption         =   "Car Exterior Finish "
      Height          =   1455
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   3135
      Begin VB.OptionButton Option3
         Caption         =   "Customized Detailing ($599.99)"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   2655
      End
      Begin VB.OptionButton Option2
         Caption         =   "Pearlized ($345.72)"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton Option1
         Caption         =   "Standard (No extra charge)"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1
      Caption         =   "Accessories"
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   3135
      Begin VB.CheckBox Check3
         Caption         =   "Computer Navigation ($1,741.43)"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   2655
      End
      Begin VB.CheckBox Check2
         Caption         =   "Leather Interior ($987.41)"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox Check1
         Caption         =   "Stereo System ($425.76)"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Label Label7
      Caption         =   "                                                 Total"
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label6
      Caption         =   "                                   Sales Tax (8%)"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label5
      Caption         =   "                                            Subtotal"
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label4
      Caption         =   "                     Accessories and Finish "
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label3
      Caption         =   "                              Car's Sales Price"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label2
      Caption         =   "                     Enter your car model"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1
      Caption         =   $"Purchasepage.frx":0000
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   8415
   End
End
Attribute VB_Name = "Purchasepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdcalculate_Click()
   Dim i As Integer = 1, p As Single, af As Single, sum As Single, s As Single, t As Single
   Dim found As Boolean

   carprice.Cls
   AccessoriesFinish.Cls
   subtotal.Cls
   Salestax.Cls
   Total.Cls

   found = False
   Do
       If UCase(carmodeltext) = carmodels(i) Then
          found = True
          p = price(i)
       End If
   Loop While i <= ctr

   If (Not found) Then
       MsgBox ("You did not enter the correct car model! Our car model is in the parentheses")
   Else
       carprice.Print FormatNumber(p, 2)

     If (Check1 = 1) Then
       af = af + 425.76
     End If
     If (Check2 = 1) Then
        af = af + 987.41
     End If
     If (Check3 = 1) Then
        af = af + 1741.43
     End If

     If Option1 Then
        af = af + 0
     ElseIf Option2 Then
        af = af + 345.72
     ElseIf Option3 Then
        af = af + 599.99
     End If

     AccessoriesFinish.Print FormatNumber(af, 2)

     s = p + af
     subtotal.Print FormatNumber(s, 2)

     t = 0.08 * s
     Salestax.Print FormatNumber(t, 2)

     sum = s + t
     Total.Print FormatNumber(sum, 2)

    End If

End Sub

Private Sub Command2_Click()
   Purchasepage.Hide
   generalpage.Show
End Sub
Private Sub cmdbonus_Click()
   Purchasepage.Hide
   Bonuspage.Show
End Sub
