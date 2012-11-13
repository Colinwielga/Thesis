VERSION 5.00
Begin VB.Form frmSubscriptionCheck 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Checking Subscription Information"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMainPage 
      Caption         =   "Go to Main"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   2
      Top             =   6240
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   4695
      Left            =   5640
      ScaleHeight     =   4635
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   1440
      Width           =   5895
   End
   Begin VB.Label lblThankYou 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Thank You, Enjoy Your Subscription!"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   720
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   6630
      Left            =   0
      Picture         =   "frmSubscriptionCheck.frx":0000
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "frmSubscriptionCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdMainPage_Click()
    frmSubscriptionCheck.Hide       'Hides SubcriptionCheck form
    frmSelectWant.Show              'Shows SelectWant form
End Sub
Private Sub Image1_Click()
    picResults.Cls
    picResults.Print "Name", , Name1                                            'prints Name and then what user entered into text box on Subscription form
    picResults.Print "Street Address", Address                                  'prints Street Address and then what user entered into text box on Subscription form
    picResults.Print "City", , City                                             'prints City and then what user entered into text box on Subscription form
    picResults.Print "State", , State                                           'prints State and then what user entered into text box on Subscription form
    picResults.Print "Postal Code", , Postal                                    'prints Postal Code and then what user entered into text box on Subscription form
    picResults.Print "Country", , Country                                       'prints Country and then what user entered into text box on Subscription form
    picResults.Print "Phone Number", "("; Area; ")"; FirstNum; "-"; LastNum     'prints Phone Number and then what user entered into text box on Subscription form
    picResults.Print "Payment Method", PaymentMethod                            'prints Payment Method and then what user entered into text box on Subscription form
    picResults.Print "Credit Card", , CardType                                  'prints Credit Card and then what user entered into text box on Subscription form
    picResults.Print "Credit Card Number", CardNumber                           'prints Credit Card Number and then what user entered into text box on Subscription form
    picResults.Print "Expiration Date", Exp1; "/"; Exp2                         'prints Expiration Date and then what user entered into text box on Subscription form
    picResults.Print "Subscription Type", SubscriptionType                      'prints SubscriptionType and then what user entered into text box on Subscription form
End Sub
