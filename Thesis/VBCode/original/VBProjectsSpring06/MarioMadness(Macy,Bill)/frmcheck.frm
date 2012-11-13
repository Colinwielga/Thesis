VERSION 5.00
Begin VB.Form frmcheck 
   Caption         =   "Verify Your Information"
   ClientHeight    =   9390
   ClientLeft      =   3690
   ClientTop       =   3105
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   Picture         =   "frmcheck.frx":0000
   ScaleHeight     =   9390
   ScaleWidth      =   10455
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change an Error"
      Height          =   1215
      Left            =   2520
      TabIndex        =   3
      Top             =   5640
      Width           =   2412
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "View Your Information"
      Height          =   1215
      Left            =   480
      TabIndex        =   2
      Top             =   4200
      Width           =   2412
   End
   Begin VB.CommandButton cmdfinalize 
      Caption         =   "Click to Finalize Your Information"
      Height          =   1215
      Left            =   5880
      TabIndex        =   1
      Top             =   840
      Width           =   2412
   End
   Begin VB.PictureBox picresults 
      Height          =   3492
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "By Bill Macy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmcheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Mario Madness
'Form name: frmCheck
'Author: Bill Macy
'Date Written: Tuesday March 14th, 2006
'Objective of form:  This form allows the user to verify the information they entered in the magazine form
                'They can either move forward and submit the form, go back and change an error to resubmit, or
                'view the information they had entered.
                
                
Option Explicit

Private Sub cmdChange_Click()
    frmMagazine.Show        'shows the magazine form
    frmcheck.Hide       'hides the check form
End Sub

Private Sub cmdfinalize_Click()
    MsgBox "Thank you for your subscription.  The first magazine should be sent to you starting the month you specified.  We appreciate your business.", , "Subscription Complete"      'thanks the user for submitting a subscription
    frmMain.Show        'shows the main form
    frmcheck.Hide       'hides the check form
    frmMagazine.Hide        'hides the magazine form
End Sub

Private Sub cmdview_Click()
    picresults.Cls      'clears the pictures box of anything that may have previoulsy been written in it
    picresults.Print "Name", , name1        'prints "Name" and then what the user entered in that particular text box on the magazine form
    picresults.Print "Month to Start Subscription", Date1; "/"; date2       'prints "Month to Start Subscription" and then what the user entered in that particular text box on the magazine form
    picresults.Print "Street Address", streetaddress        'prints "Street Address" and then what the user entered in that particular text box on the magazine form
    picresults.Print "P.O. Box Number", pobox       'prints "P.O. Box Number" and then what the user entered in that particular text box on the magazine form
    picresults.Print "City", , city     'prints "City" and then what the user entered in that particular text box on the magazine form
    picresults.Print "State", , state       'prints "State" and then what the user entered in that particular text box on the magazine form
    picresults.Print "Postal Code", , postalcode      'prints "Postal Code" and then what the user entered in that particular text box on the magazine form
    picresults.Print "Country", , country       'prints "Country" and then what the user entered in that particular text box on the magazine form
    picresults.Print "Email Address", emailaddress      'prints "Email Address" and then what the user entered in that particular text box on the magazine form
    picresults.Print "Phone Number", "("; phonenumber; ")"; phonenumber2; "-"; phonenumber3     'prints "Phone Number" and then what the user entered in that particular text box on the magazine form
    picresults.Print "Payment Type", paymenttype       'prints "Payment Type" and then what the user entered in that particular text box on the magazine form
    picresults.Print "Credit Card Name", cardname       'prints "Credit Card Name" and then what the user entered in that particular text box on the magazine form
    picresults.Print "Card Type", , cardtype        'prints "Card Type" and then what the user entered in that particular text box on the magazine form
    picresults.Print "Credit Card Number", cardnumber       'prints "Credit Card Number" and then what the user entered in that particular text box on the magazine form
    picresults.Print "Expiration Date", expdate; "/"; expdate2      'prints "Expiration Date" and then what the user entered in that particular text box on the magazine form
    picresults.Print "Subscription Type", subscriptiontype      'prints "Subscription Type" and then what the user entered in that particular text box on the magazine form
End Sub
    
Private Sub Form_Load()
    picresults.Cls      'clears the picture box of anything that may have previously been written in it
End Sub
