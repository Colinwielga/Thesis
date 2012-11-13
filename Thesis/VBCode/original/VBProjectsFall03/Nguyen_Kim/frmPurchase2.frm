VERSION 5.00
Begin VB.Form frmPurchase 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Kim Nguyen"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   7920
      MaskColor       =   &H00FF8080&
      TabIndex        =   15
      Top             =   7200
      Width           =   2415
   End
   Begin VB.CommandButton cmdHomepage 
      Caption         =   "Return to Homepage"
      Height          =   735
      Left            =   2880
      MaskColor       =   &H00FF8080&
      TabIndex        =   14
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdLoan 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Loan Payment Calculator"
      Height          =   735
      Left            =   240
      MaskColor       =   &H00FF8080&
      TabIndex        =   13
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF0000&
      Caption         =   "Price Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   6375
      Left            =   6360
      TabIndex        =   11
      Top             =   360
      Width           =   3855
      Begin VB.PictureBox picPriceInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   240
         ScaleHeight     =   5355
         ScaleWidth      =   3315
         TabIndex        =   12
         Top             =   480
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Exterior Finish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1455
      Left            =   3000
      TabIndex        =   7
      Top             =   4920
      Width           =   2415
      Begin VB.OptionButton OptCustomized 
         Caption         =   "Customized Detailing"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2055
      End
      Begin VB.OptionButton OptPearlized 
         Caption         =   "Pearlized"
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton OptStandard 
         Caption         =   "Standard"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Accessories"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   2295
      Begin VB.CheckBox checkcomputer 
         Caption         =   "Computer Nevagation"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox checkLeather 
         Caption         =   "Leather Interior"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox checkStereo 
         Caption         =   "Stereo System"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.PictureBox picCarPurchase 
      Height          =   2415
      Left            =   240
      ScaleHeight     =   2355
      ScaleWidth      =   4755
      TabIndex        =   2
      Top             =   1440
      Width           =   4815
   End
   Begin VB.ComboBox cbo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      ItemData        =   "frmPurchase2.frx":0000
      Left            =   120
      List            =   "frmPurchase2.frx":0022
      TabIndex        =   1
      Text            =   "Select A Car That You Want To Buy"
      Top             =   360
      Width           =   5055
   End
   Begin VB.CommandButton cmdTotalPrice 
      Caption         =   "Total Price of The Car"
      Height          =   735
      Left            =   5520
      TabIndex        =   0
      Top             =   7200
      Width           =   2055
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Online Car Dealer
'Form Name : frmPurchase (FrmPurchase.frm)
'Author: Kim Nguyen
'Date Written: October 29, 2003
'Purpose of Form: Let the user pick out a car that they want to by in
                'the combo drop down list, which the picture will pop up
                'when the car name get selected
                'then the price will be print out in the
                'if the user want to add accessory or finishes they can too and the
                'cost of each item will be add to the runing total with tax included
                

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.Option Explicit
Option Explicit
'Declares the variables

Dim Year(1 To 10) As Integer
Dim CarName(1 To 10) As String
Dim Price(1 To 10) As Long
Dim Class(1 To 10) As String
Dim Cylinders(1 To 10) As String
Dim Horsepower(1 To 10) As Single
Dim StandardTransmission(1 To 10) As String
Dim Drivetrain(1 To 10) As String
Dim PicName(1 To 10) As String
Dim Finish As Single
Dim Accessory As Single
Dim Total1 As Single
Dim Total2 As Single
Dim Total3 As Single
Dim Subtotal As Single
Dim TotalPrice As Single
Dim Tax As Single
Dim I As Integer

'Show the Loan payment calculator
Private Sub cmdLoan_Click()
frmLoan.Show
End Sub
'load the file and put it into an array
Public Sub Form_Load()
Open Path & "Cars.txt" For Input As #3
For I = 1 To 10
    Input #3, Year(I), CarName(I), Price(I), Class(I), Cylinders(I), Horsepower(I), StandardTransmission(I), Drivetrain(I), PicName(I)
Next I
Close #1
End Sub
'
Private Sub cbo2_Click()
    picPriceInfo.Cls
'I is the counter, so everytime a user click on the combo dropdown
'and click on the name of the car on the Indexlist, the position on the indexlist
'starts from 0 that's why I need to add 1 so that it would match with the array in the file
I = cbo2.ListIndex + 1

'I indicate the position that the CarName, Price and Picture of the car will be printed
    picCarPurchase.Picture = LoadPicture(Path & PicName(I))
    picPriceInfo.Print CarName(I)
    picPriceInfo.Print
    picPriceInfo.Print "Car Sales Price"; Tab(20); FormatCurrency(Price(I), 2)
    picPriceInfo.Print
    
End Sub
'425.76 will be added to the total if this box is checked
Private Sub checkStereo_Click()
Total1 = 425.76
picPriceInfo.Print "Stereo System"; Tab(20); FormatCurrency(425.76, 2)
End Sub
'987.41 will be added to the total if this box is checked
Private Sub checkLeather_Click()
Total2 = 987.41
picPriceInfo.Print "Leather Interior"; Tab(20); FormatCurrency(987.41, 2)
End Sub
'1741.23 will be added to the total if this box is checked
Private Sub checkcomputer_Click()
Total3 = 1741.23
picPriceInfo.Print "Cumputer Nevigation"; Tab(20); FormatCurrency(1741.23, 2)
End Sub

'If one of the option box is checked than the cost of that option will be add to the
 'running total
 'TotalPrice of the car include the sales price plus accecessory or finish if there is
 'any and tax will also be included
 
Private Sub cmdTotalPrice_Click()
    picPriceInfo.Print
If OptStandard = True Then
    picPriceInfo.Print "Finish"; Tab(20); "$ 0.0"
    Finish = 0
ElseIf OptPearlized = True Then
    picPriceInfo.Print "Finish"; Tab(20); "$345.72"
    Finish = 345.72
ElseIf OptCustomized = True Then
    picPriceInfo.Print "Finish"; Tab(20); "$599.99"
    Finish = 599.99
End If
    Subtotal = Price(I) + Total1 + Total2 + Total3 + Finish
picPriceInfo.Print
picPriceInfo.Print "Subtotal Price"; Tab(20); FormatCurrency(Subtotal, 2)
    Tax = Subtotal * 0.08
picPriceInfo.Print "Tax"; Tab(20); FormatCurrency(Tax, 2)
picPriceInfo.Print
TotalPrice = Subtotal + Tax
picPriceInfo.Print "Total Price"; Tab(20); FormatCurrency(TotalPrice, 2)



End Sub
'show the Info form and hide this Purchase form
Private Sub cmdHomepage_Click()
frmInfo.Show
frmPurchase.Hide
End Sub
'End the program now
Private Sub cmdQuit_Click()
End
End Sub


